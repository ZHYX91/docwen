from __future__ import annotations

import argparse
import ast
import importlib.util
import json
import os
import re
import stat
import subprocess
import sys
import tempfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:
    _BOOTSTRAP_ROOT = Path(__file__).resolve().parents[1]
    if str(_BOOTSTRAP_ROOT) not in sys.path:
        sys.path.insert(0, str(_BOOTSTRAP_ROOT))

from tools.windows_short_path import ShortPathDriveError, drive_root, mount_short_drive, unmount_short_drive
from tools.workspace_root import WORKSPACE_ROOT_ENV as _WORKSPACE_ROOT_ENV
from tools.workspace_root import WorkspaceRootError, resolve_workspace_root

FAST_MARK_EXPR = "(unit or contract) and not slow"
PR_GATE_MARK_EXPR = "pr_gate and (integration or gui_smoke or e2e)"
RELEASE_GATE_MARK_EXPR = "release_gate and (integration or gui_smoke or e2e)"
PYTEST_BASE_ADDOPTS = "-v --tb=short --strict-markers --import-mode=importlib -ra"
PYTEST_XDIST_ENV = "DOCWEN_PYTEST_XDIST"
PYTEST_XDIST_WORKERS_ENV = "DOCWEN_PYTEST_XDIST_WORKERS"
PYTEST_XDIST_MAX_PROCESSES_ENV = "DOCWEN_PYTEST_XDIST_MAX_PROCESSES"
PYTEST_XDIST_DEFAULT_MAX_PROCESSES = "6"
PYTEST_XDIST_DIST = "loadfile"
PYTEST_PRIMARY_MARKER_DEBT_LIMIT = 0
PYTEST_PRIMARY_MARKER_OVERLAP_LIMIT = 0
PYTEST_RUNTIME_ROOT_ENV = "DOCWEN_PYTEST_RUNTIME_ROOT"
PYTEST_REPORT_DIR_ENV = "DOCWEN_PYTEST_REPORT_DIR"
WORKSPACE_ROOT_ENV = _WORKSPACE_ROOT_ENV
PYTEST_RUNTIME_LEASE = ".docwen-temp-lease.json"
PYTEST_RUNTIME_PREFIX = "p"
PYTEST_BASETEMP_NAME = "b"
PYTEST_CACHE_NAME = "c"
PYTEST_REPORT_NAME = "reports"
PYTEST_SYSTEM_TEMP_NAME = "t"


def _run(args: list[str], *, env: dict[str, str] | None = None) -> int:
    proc = subprocess.run(args, cwd=Path(__file__).resolve().parents[1], env=env)
    return int(proc.returncode)


def _path_traverses_link_or_reparse(path: Path, *, stop_at: Path | None = None) -> bool:
    current = path.absolute()
    boundary = stop_at.absolute() if stop_at is not None else None
    while True:
        try:
            metadata = current.lstat()
        except FileNotFoundError:
            pass
        except OSError as error:
            raise ValueError(f"cannot inspect pytest runtime path {current}: {error}") from error
        else:
            file_attributes = int(getattr(metadata, "st_file_attributes", 0))
            reparse_flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
            if stat.S_ISLNK(metadata.st_mode) or (reparse_flag and file_attributes & reparse_flag):
                return True
        if current == boundary:
            return False
        parent = current.parent
        if parent == current:
            return False
        current = parent


def _write_runtime_lease(runtime_root: Path, *, state: str) -> None:
    marker = runtime_root / PYTEST_RUNTIME_LEASE
    payload = {
        "schemaVersion": 1,
        "owner": "docwen.tools.qa",
        "kind": "pytest-runtime",
        "pid": os.getpid(),
        "createdAt": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "state": state,
        "root": str(runtime_root),
    }
    marker.write_text(
        json.dumps(payload, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )


def _patch_runtime_lease(
    runtime_root: Path,
    *,
    state: str | None = None,
    fields: dict[str, Any] | None = None,
    remove_fields: tuple[str, ...] = (),
) -> None:
    marker = runtime_root / PYTEST_RUNTIME_LEASE
    payload = json.loads(marker.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"pytest runtime lease is not an object: {runtime_root}")
    if not str(payload.get("owner", "")).startswith("docwen.") or payload.get("root") != str(runtime_root):
        raise ValueError(f"pytest runtime lease mismatch: {runtime_root}")
    if state is not None:
        payload["state"] = state
    if fields:
        payload.update(fields)
    for field in remove_fields:
        payload.pop(field, None)
    replacement = marker.with_suffix(".tmp")
    replacement.write_text(
        json.dumps(payload, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    replacement.replace(marker)


def _update_runtime_lease(runtime_root: Path, *, state: str) -> None:
    _patch_runtime_lease(runtime_root, state=state)


def _runtime_has_managed_lease(runtime_root: Path) -> bool:
    marker = runtime_root / PYTEST_RUNTIME_LEASE
    try:
        payload = json.loads(marker.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    return (
        isinstance(payload, dict)
        and str(payload.get("owner", "")).startswith("docwen.")
        and payload.get("root") == str(runtime_root)
    )


def _windows_extended_path(path: Path) -> str:
    absolute = os.path.abspath(os.fspath(path))
    if os.name != "nt" or absolute.startswith("\\\\?\\"):
        return absolute
    if absolute.startswith("\\\\"):
        return f"\\\\?\\UNC\\{absolute[2:]}"
    return f"\\\\?\\{absolute}"


def _remove_owned_readonly_path(function: object, path: str, error: BaseException) -> None:
    if not isinstance(error, PermissionError) or not callable(function):
        raise error
    os.chmod(path, stat.S_IWRITE)
    function(path)


def _cleanup_owned_runtime(runtime_root: Path) -> None:
    import shutil

    marker = runtime_root / PYTEST_RUNTIME_LEASE
    payload = json.loads(marker.read_text(encoding="utf-8"))
    if payload.get("owner") != "docwen.tools.qa" or payload.get("root") != str(runtime_root):
        raise ValueError(f"pytest runtime lease mismatch: {runtime_root}")
    if _path_traverses_link_or_reparse(runtime_root):
        raise ValueError(f"pytest runtime cleanup refuses a link or reparse point: {runtime_root}")
    shutil.rmtree(_windows_extended_path(runtime_root), onexc=_remove_owned_readonly_path)


def _default_pytest_parent(repo_root: Path, workspace_root: Path | None = None) -> Path:
    workspace_root = resolve_workspace_root(repo_root, explicit=workspace_root)
    return workspace_root / "temp"


def _optional_workspace_root(repo_root: Path, workspace_root: Path | None = None) -> Path | None:
    try:
        return resolve_workspace_root(repo_root, explicit=workspace_root)
    except WorkspaceRootError:
        return None


def _is_within(path: Path, parent: Path) -> bool:
    return path == parent or parent in path.parents


def _workspace_runtime_parent(workspace_root: Path) -> Path:
    return workspace_root / "temp"


def _cleanup_expired_workspace_temps(workspace_root: Path) -> None:
    from tools import workspace_cleanup

    plan = workspace_cleanup.create_plan(workspace_root=workspace_root)
    entries = plan.get("entries", [])
    if not entries:
        return

    timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
    plan_path = workspace_root / "diagnostics" / f"qa-housekeeping-{os.getpid()}-{timestamp}.json"
    workspace_cleanup.save_plan(plan, plan_path)
    workspace_cleanup.apply_saved_plan(plan_path, workspace_root=workspace_root)
    plan_path.unlink()
    print(f"[qa] removed {len(entries)} workspace scratch root(s) via the shared retention plan")


def _pytest_runtime_environment(
    repo_root: Path,
    requested_root: Path | None,
    *,
    own_requested_root: bool = False,
    workspace_root: Path | None = None,
) -> tuple[Path, dict[str, str], bool]:
    configured = requested_root
    if configured is None:
        configured_value = os.environ.get(PYTEST_RUNTIME_ROOT_ENV, "").strip()
        configured = Path(configured_value) if configured_value else None
    governed_workspace = _optional_workspace_root(repo_root, workspace_root)
    owned = configured is None or own_requested_root
    if configured is None:
        if governed_workspace is None:
            raise ValueError(f"governed workspace root is required for pytest runtime storage: {repo_root}")
        runtime_parent = _default_pytest_parent(repo_root, governed_workspace).absolute()
        if _path_traverses_link_or_reparse(runtime_parent):
            raise ValueError(f"pytest runtime parent must not traverse a link or reparse point: {runtime_parent}")
        runtime_parent.mkdir(parents=True, exist_ok=True)
        runtime_candidate = Path(tempfile.mkdtemp(prefix=PYTEST_RUNTIME_PREFIX, dir=runtime_parent)).absolute()
    else:
        runtime_candidate = configured.absolute()
        if own_requested_root and runtime_candidate.exists():
            raise ValueError(f"owned pytest runtime root must not already exist: {runtime_candidate}")
    if _path_traverses_link_or_reparse(runtime_candidate):
        raise ValueError(f"pytest runtime root must not traverse a link or reparse point: {runtime_candidate}")
    prospective_root = runtime_candidate.resolve(strict=False)
    if prospective_root == repo_root or repo_root in prospective_root.parents:
        raise ValueError(f"pytest runtime root must be outside the repository: {prospective_root}")
    if governed_workspace is not None:
        governed_temp = _workspace_runtime_parent(governed_workspace).resolve(strict=False)
        if not _is_within(prospective_root, governed_temp):
            raise ValueError(
                "pytest runtime root must stay inside the governed workspace temp tree: "
                f"{prospective_root} not under {governed_temp}"
            )
    if runtime_candidate.exists() and not runtime_candidate.is_dir():
        raise ValueError(f"pytest runtime root is not a directory: {runtime_candidate}")
    runtime_candidate.mkdir(parents=True, exist_ok=True)
    runtime_root = runtime_candidate.resolve()
    basetemp = runtime_root / PYTEST_BASETEMP_NAME
    cache_dir = runtime_root / PYTEST_CACHE_NAME
    report_dir = runtime_root / PYTEST_REPORT_NAME
    system_temp = runtime_root / PYTEST_SYSTEM_TEMP_NAME
    for path in (basetemp, cache_dir, report_dir, system_temp):
        path.mkdir(parents=True, exist_ok=True)
        if _path_traverses_link_or_reparse(path, stop_at=runtime_root):
            raise ValueError(f"pytest runtime child must not traverse a link or reparse point: {path}")
    if owned:
        _write_runtime_lease(runtime_root, state="active")
    environment = os.environ.copy()
    environment.update(
        {
            PYTEST_RUNTIME_ROOT_ENV: str(runtime_root),
            PYTEST_REPORT_DIR_ENV: str(report_dir),
            "DOCWEN_PYTEST_MAX_MISSING_PRIMARY_MARKERS": str(PYTEST_PRIMARY_MARKER_DEBT_LIMIT),
            "PYTHONDONTWRITEBYTECODE": "1",
            "TEMP": str(system_temp),
            "TMP": str(system_temp),
            "TMPDIR": str(system_temp),
        }
    )
    return runtime_root, environment, owned


def _runtime_environment_for_view(environment: dict[str, str], runtime_view: Path) -> dict[str, str]:
    projected = environment.copy()
    projected.update(
        {
            PYTEST_RUNTIME_ROOT_ENV: str(runtime_view),
            PYTEST_REPORT_DIR_ENV: str(runtime_view / PYTEST_REPORT_NAME),
            "TEMP": str(runtime_view / PYTEST_SYSTEM_TEMP_NAME),
            "TMP": str(runtime_view / PYTEST_SYSTEM_TEMP_NAME),
            "TMPDIR": str(runtime_view / PYTEST_SYSTEM_TEMP_NAME),
        }
    )
    return projected


def _should_use_short_runtime_drive(*, suite: str) -> bool:
    return os.name == "nt" and suite == "full"


def _pytest_base_cmd() -> list[str]:
    return [sys.executable, "-m", "pytest", *_xdist_args(), "-o", f"addopts={PYTEST_BASE_ADDOPTS}"]


def _xdist_args() -> list[str]:
    setting = os.environ.get(PYTEST_XDIST_ENV, "1").strip().lower() or "1"
    if setting in {"0", "false", "no", "off"}:
        return []
    if setting not in {"1", "true", "yes", "on"}:
        raise ValueError(f"{PYTEST_XDIST_ENV} must be a boolean value")
    if importlib.util.find_spec("xdist") is None:
        print(
            f"[qa] {PYTEST_XDIST_ENV}=1 but pytest-xdist is unavailable; falling back to single-process pytest",
            file=sys.stderr,
        )
        return []
    workers = os.environ.get(PYTEST_XDIST_WORKERS_ENV, "auto").strip() or "auto"
    args = ["-n", workers, "--dist", PYTEST_XDIST_DIST]
    if workers in {"auto", "logical"}:
        max_processes = (
            os.environ.get(PYTEST_XDIST_MAX_PROCESSES_ENV, PYTEST_XDIST_DEFAULT_MAX_PROCESSES).strip()
            or PYTEST_XDIST_DEFAULT_MAX_PROCESSES
        )
        if not max_processes.isdecimal() or int(max_processes) < 1:
            raise ValueError(f"{PYTEST_XDIST_MAX_PROCESSES_ENV} must be a positive integer")
        args += ["--maxprocesses", max_processes]
    return args


def _platform_fast_mark_expr() -> str:
    expr_parts = [FAST_MARK_EXPR]
    if sys.platform != "win32":
        expr_parts.append("not windows_only")
    if sys.platform != "darwin":
        expr_parts.append("not macos_only")
    if not sys.platform.startswith("linux"):
        expr_parts.append("not linux_only")
    return " and ".join(expr_parts)


def _platform_release_mark_expr() -> str:
    expr_parts = [RELEASE_GATE_MARK_EXPR]
    if sys.platform != "win32":
        expr_parts.append("not windows_only")
    if sys.platform != "darwin":
        expr_parts.append("not macos_only")
    if not sys.platform.startswith("linux"):
        expr_parts.append("not linux_only")
    return " and ".join(expr_parts)


def _scan_private_symbol_usage(repo_root: Path) -> int:
    deny_import = re.compile(r"^\s*from\s+docwen(?:\.[\w]+)*\s+import\s+.*\b_\w+", re.MULTILINE)
    deny_string_path = re.compile(r"docwen(?:\.[\w]+)*\._[A-Za-z]\w*")

    def _chain(node: ast.AST) -> list[str] | None:
        if isinstance(node, ast.Name):
            return [node.id]
        if isinstance(node, ast.Attribute):
            parent = _chain(node.value)
            if not parent:
                return None
            return [*parent, node.attr]
        return None

    def _scan_ast_for_docwen_private_access(text: str) -> list[tuple[int, str]]:
        try:
            tree = ast.parse(text)
        except SyntaxError:
            return []

        aliases: dict[str, str] = {}
        issues: list[tuple[int, str]] = []

        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for item in node.names:
                    if (item.name == "docwen" or item.name.startswith("docwen.")) and item.asname:
                        aliases[item.asname] = item.name
            elif isinstance(node, ast.ImportFrom):
                mod = node.module or ""
                if not (mod == "docwen" or mod.startswith("docwen.")):
                    continue
                for item in node.names:
                    if item.name == "*":
                        continue
                    if item.name.startswith("_") and not item.name.startswith("__"):
                        issues.append((getattr(node, "lineno", 1), "import-private-from-docwen"))
                        continue
                    bound = item.asname or item.name
                    aliases[bound] = f"{mod}.{item.name}"
            elif isinstance(node, ast.Attribute):
                if not (node.attr.startswith("_") and not node.attr.startswith("__")):
                    continue

                dotted = _chain(node)
                if not dotted:
                    continue
                root = dotted[0]
                if root == "docwen":
                    issues.append((getattr(node, "lineno", 1), "attr-private-on-docwen"))
                    continue
                if root in aliases:
                    issues.append((getattr(node, "lineno", 1), "attr-private-on-docwen-import"))
                    continue

        return issues

    offenders: dict[Path, set[str]] = {}
    for rel_root in ("tests", "tools"):
        base = repo_root / rel_root
        if not base.is_dir():
            continue
        for path in base.rglob("*.py"):
            if path.name == "qa.py":
                continue
            try:
                text = path.read_text(encoding="utf-8")
            except Exception:
                continue

            if deny_import.search(text):
                offenders.setdefault(path, set()).add("import-private-from-docwen")
            if deny_string_path.search(text):
                offenders.setdefault(path, set()).add("string-refers-docwen-private")

            for lineno, kind in _scan_ast_for_docwen_private_access(text):
                offenders.setdefault(path, set()).add(f"{kind}:L{lineno}")

    if not offenders:
        return 0

    print("==> private-symbol-boundary")
    print("Found disallowed private symbol usage in tests/tools. Fix by adding a public API in src and importing that.")
    for path, reasons in sorted(offenders.items(), key=lambda x: str(x[0])):
        rel = path.relative_to(repo_root)
        print(f"  - {rel} ({','.join(sorted(reasons))})")
    return 1


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--suite", choices=["fast", "full", "pr-integration", "release"], default="fast")
    parser.add_argument("--phase5", action="store_true")
    parser.add_argument(
        "--workspace-root",
        type=Path,
        help="Use an explicit DocWen governance root for default QA runtime storage.",
    )
    parser.add_argument("--skip-ruff", action="store_true")
    parser.add_argument("--skip-pyright", action="store_true")
    parser.add_argument("--skip-pytest", action="store_true")
    parser.add_argument(
        "--pytest-basetemp",
        type=Path,
        help="Use an explicit pytest temporary root instead of the machine-global default.",
    )
    parser.add_argument(
        "--pytest-runtime-root",
        type=Path,
        help="Use one root inside the governed workspace temp tree for pytest runtime data.",
    )
    parser.add_argument(
        "--keep-pytest-runtime",
        action="store_true",
        help="Keep an automatically created pytest runtime after a successful run.",
    )
    parser.add_argument(
        "--own-pytest-runtime",
        action="store_true",
        help="Own and clean a new explicit workspace-contained runtime root; rejects a pre-existing path.",
    )
    args = parser.parse_args(argv)
    if args.own_pytest_runtime and args.pytest_runtime_root is None and args.pytest_basetemp is None:
        parser.error("--own-pytest-runtime requires --pytest-runtime-root or --pytest-basetemp")

    repo_root = Path(__file__).resolve().parents[1]
    code = _scan_private_symbol_usage(repo_root)
    if code != 0:
        return code

    steps: list[tuple[str, list[str], bool]] = []
    if not args.skip_ruff:
        steps += [
            ("ruff-format", [sys.executable, "-m", "ruff", "format", "--check", "."], True),
            ("ruff-check", [sys.executable, "-m", "ruff", "check", "."], True),
        ]
        if args.phase5:
            steps += [
                (
                    "ruff-phase5",
                    [
                        sys.executable,
                        "-m",
                        "ruff",
                        "check",
                        ".",
                        "--select",
                        "PTH,T20",
                        "--statistics",
                        "--exit-zero",
                    ],
                    False,
                )
            ]
    if not args.skip_pyright:
        steps += [
            (
                "pyright",
                [sys.executable, "-m", "pyright", "--level", "error", "--pythonpath", sys.executable],
                True,
            ),
        ]

    exit_code = 0
    for name, cmd, gate in steps:
        print(f"==> {name}")
        code = _run(cmd)
        if code != 0:
            if gate:
                return code
            exit_code = code

    if not args.skip_pytest:
        print("==> pytest")
        selected_workspace = _optional_workspace_root(repo_root, args.workspace_root)
        if selected_workspace is not None:
            try:
                _cleanup_expired_workspace_temps(selected_workspace)
            except (OSError, ValueError) as error:
                print(f"[qa] workspace runtime cleanup failed: {error}", file=sys.stderr)
                return 2
        requested_runtime_root = args.pytest_runtime_root
        if requested_runtime_root is None and args.pytest_basetemp is not None:
            requested_runtime_root = args.pytest_basetemp
        try:
            runtime_root, pytest_environment, runtime_owned = _pytest_runtime_environment(
                repo_root,
                requested_runtime_root,
                own_requested_root=args.own_pytest_runtime,
                workspace_root=selected_workspace,
            )
        except ValueError as error:
            print(f"[qa] {error}", file=sys.stderr)
            return 2
        print(f"[qa] pytest physical runtime root: {runtime_root}")

        short_drive: str | None = None
        runtime_view = runtime_root
        lease_managed = _runtime_has_managed_lease(runtime_root)
        return_code: int | None = None
        interrupted = False
        unexpected_error: BaseException | None = None
        short_drive_error: ShortPathDriveError | None = None
        try:
            if _should_use_short_runtime_drive(suite=args.suite):
                short_drive = mount_short_drive(runtime_root)
                runtime_view = drive_root(short_drive)
                if lease_managed:
                    _patch_runtime_lease(runtime_root, fields={"shortDrive": short_drive})
                print(f"[qa] pytest short runtime view: {runtime_view}")

            pytest_cmd = _pytest_base_cmd()
            pytest_cmd += [
                "--basetemp",
                str(runtime_view / PYTEST_BASETEMP_NAME),
                "-o",
                f"cache_dir={runtime_view / PYTEST_CACHE_NAME}",
            ]
            if args.suite == "fast":
                pytest_cmd += ["-m", _platform_fast_mark_expr()]
            elif args.suite == "pr-integration":
                pytest_cmd += ["-m", PR_GATE_MARK_EXPR]
            elif args.suite == "release":
                pytest_cmd += ["-m", _platform_release_mark_expr()]
            return_code = _run(
                pytest_cmd,
                env=_runtime_environment_for_view(pytest_environment, runtime_view),
            )
        except KeyboardInterrupt:
            interrupted = True
        except BaseException as error:
            unexpected_error = error
        finally:
            if short_drive is not None:
                try:
                    unmount_short_drive(short_drive, expected_target=runtime_root)
                except ShortPathDriveError as error:
                    short_drive_error = error
                else:
                    if lease_managed:
                        _patch_runtime_lease(runtime_root, remove_fields=("shortDrive",))

        if interrupted:
            if lease_managed:
                _update_runtime_lease(runtime_root, state="retained-interrupted")
            print(f"[qa] interrupted pytest runtime retained: {runtime_root}", file=sys.stderr)
            if short_drive_error is not None:
                print(f"[qa] short runtime drive cleanup failed: {short_drive_error}", file=sys.stderr)
            return 130
        if unexpected_error is not None:
            if lease_managed:
                _update_runtime_lease(runtime_root, state="retained-interrupted")
            if short_drive_error is not None:
                unexpected_error.add_note(f"short runtime drive cleanup failed: {short_drive_error}")
            raise unexpected_error
        if short_drive_error is not None:
            if lease_managed:
                _update_runtime_lease(runtime_root, state="retained-cleanup-failure")
            print(f"[qa] short runtime drive cleanup failed: {short_drive_error}", file=sys.stderr)
            return 2
        if return_code is None:
            raise RuntimeError("pytest runner returned no status")
        if return_code != 0:
            if runtime_owned:
                _update_runtime_lease(runtime_root, state="retained-failure")
                print(f"[qa] failed pytest runtime retained: {runtime_root}", file=sys.stderr)
            return return_code
        if runtime_owned:
            if args.keep_pytest_runtime:
                _update_runtime_lease(runtime_root, state="retained-manual")
                print(f"[qa] pytest runtime retained by request: {runtime_root}")
            else:
                try:
                    _cleanup_owned_runtime(runtime_root)
                except (OSError, ValueError) as error:
                    print(f"[qa] pytest runtime cleanup failed: {error}", file=sys.stderr)
                    return 2
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
