from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from xml.etree import ElementTree

import tomlkit


@dataclass(frozen=True)
class CoverageDomain:
    label: str
    matches: tuple[str, ...]


@dataclass
class CoverageSummary:
    files: set[str]
    covered_lines: int = 0
    total_lines: int = 0

    @property
    def percent(self) -> float:
        if self.total_lines == 0:
            return 0.0
        return self.covered_lines * 100.0 / self.total_lines


CORE_DOMAINS: tuple[CoverageDomain, ...] = (
    CoverageDomain("core/", ("packages/core/src/docwen_core/",)),
    CoverageDomain("application/", ("packages/application/src/docwen_application/",)),
    CoverageDomain("runtime/", ("packages/runtime/src/docwen_runtime/",)),
)

# === 覆盖率爬升机制 ===
# N=3, M=2: 首版取值，调整需在 commit message 中说明原因.
CLIMBING_N = 3
CLIMBING_M = 2
CLIMBING_BUFFER_PP = 1.0

_HISTORY_PATH = Path(__file__).resolve().parent / ".coverage_history.json"

# 当前档位从 pyproject.toml [tool.coverage.report] fail_under 读取.
# 下一档位候选值: 步骤 1 测量后填入; 当前为 None 表示暂未执行测量.
CURRENT_THRESHOLD: int | None = None
NEXT_THRESHOLD: int | None = None

COMBINED_DOMAIN_LABEL = "combined"
SOFT_GATE_THRESHOLDS: tuple[tuple[str, float], ...] = (
    ("core/", 55.0),
    ("application/", 55.0),
    ("runtime/", 55.0),
    (COMBINED_DOMAIN_LABEL, 55.0),
)


def _normalize_coverage_filename(filename: str) -> str:
    normalized = filename.replace("\\", "/").strip()
    while normalized.startswith("./"):
        normalized = normalized[2:]
    return normalized


def _matches_domain(filename: str, domain: CoverageDomain) -> bool:
    normalized = _normalize_coverage_filename(filename)
    for pattern in domain.matches:
        if pattern.endswith("/"):
            stripped = pattern.rstrip("/")
            if normalized.startswith(f"{stripped}/") or f"/{stripped}/" in normalized:
                return True
        elif normalized == pattern or normalized.endswith(f"/{pattern}"):
            return True
    return False


def _summaries_from_coverage_xml(coverage_xml: Path) -> dict[str, CoverageSummary]:
    root = ElementTree.parse(coverage_xml).getroot()
    summaries = {domain.label: CoverageSummary(files=set()) for domain in CORE_DOMAINS}

    for class_node in root.findall(".//class"):
        filename = class_node.attrib.get("filename", "")
        if not filename:
            continue

        matching_labels = [domain.label for domain in CORE_DOMAINS if _matches_domain(filename, domain)]
        if not matching_labels:
            continue

        lines = class_node.findall("./lines/line")
        total_lines = len(lines)
        covered_lines = sum(1 for line in lines if int(line.attrib.get("hits", "0")) > 0)
        normalized = _normalize_coverage_filename(filename)

        for label in matching_labels:
            summary = summaries[label]
            summary.files.add(normalized)
            summary.covered_lines += covered_lines
            summary.total_lines += total_lines

    return summaries


def _parse_fail_under(items: list[str]) -> dict[str, float]:
    thresholds: dict[str, float] = {}
    for raw in items:
        name, sep, value = raw.partition("=")
        if not sep:
            raise ValueError(f"invalid --fail-under value: {raw!r}; expected <domain>=<percent>")
        thresholds[name.strip()] = float(value)
    return thresholds


def _emit_annotation(level: str, message: str) -> None:
    if os.getenv("GITHUB_ACTIONS") == "true":
        print(f"::{level}::{message}")


def _threshold_target(
    label: str, summaries: dict[str, CoverageSummary], combined_summary: CoverageSummary
) -> CoverageSummary | None:
    if label == COMBINED_DOMAIN_LABEL:
        return combined_summary
    return summaries.get(label)


def _format_threshold_failure(prefix: str, label: str, percent: float, threshold: float) -> str:
    return f"[{prefix}] {label}: {percent:.2f}% < required {threshold:.2f}%"


def _evaluate_thresholds(
    *,
    thresholds: dict[str, float],
    prefix: str,
    annotation_level: str,
    fail_on_breach: bool,
    summaries: dict[str, CoverageSummary],
    combined_summary: CoverageSummary,
) -> int:
    exit_code = 0
    for label, threshold in thresholds.items():
        summary = _threshold_target(label, summaries, combined_summary)
        if summary is None:
            print(f"[invalid-threshold] unknown domain: {label}", file=sys.stderr)
            _emit_annotation("error", f"core-domain-coverage invalid threshold target: {label}")
            exit_code = 1
            continue
        if label != COMBINED_DOMAIN_LABEL and not summary.files:
            print(f"[{prefix}] {label}: no matching files found", file=sys.stderr)
            _emit_annotation("error", f"core-domain-coverage missing domain for threshold: {label}")
            exit_code = 1
            continue
        if summary.total_lines == 0:
            continue
        if summary.percent < threshold:
            message = _format_threshold_failure(prefix, label, summary.percent, threshold)
            print(message, file=sys.stderr)
            _emit_annotation(
                annotation_level, f"core-domain-coverage {label} {summary.percent:.2f}% < {threshold:.2f}%"
            )
            if fail_on_breach:
                exit_code = 1
    return exit_code


def _print_report(summaries: dict[str, CoverageSummary]) -> CoverageSummary:
    combined_summary = CoverageSummary(files=set())

    print("==> core-domain-coverage")
    for domain in CORE_DOMAINS:
        summary = summaries[domain.label]
        if not summary.files:
            print(f"[missing] {domain.label}: no matching files found in coverage XML")
            continue

        combined_summary.files.update(summary.files)
        combined_summary.covered_lines += summary.covered_lines
        combined_summary.total_lines += summary.total_lines
        print(
            f"{domain.label}: {summary.percent:.2f}% "
            f"({summary.covered_lines}/{summary.total_lines} lines, {len(summary.files)} files)"
        )

    return combined_summary


def _read_pyproject_fail_under(repo_root: Path) -> int:
    pyproject_path = repo_root / "pyproject.toml"
    if not pyproject_path.is_file():
        raise FileNotFoundError(f"pyproject.toml not found at {pyproject_path}")
    data = tomlkit.parse(pyproject_path.read_text(encoding="utf-8"))
    return int(data["tool"]["coverage"]["report"]["fail_under"])  # type: ignore[index]


def _read_coverage_history() -> list[dict]:
    if not _HISTORY_PATH.is_file():
        return []
    try:
        with _HISTORY_PATH.open("r", encoding="utf-8") as fh:
            return json.load(fh)
    except (json.JSONDecodeError, OSError):
        return []


def _save_coverage_snapshot(snapshots: list[dict]) -> None:
    _HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    with _HISTORY_PATH.open("w", encoding="utf-8") as fh:
        json.dump(snapshots, fh, ensure_ascii=False, indent=2)


def _append_coverage_snapshot(summaries: dict[str, CoverageSummary], combined: CoverageSummary) -> None:
    record: dict = {
        "timestamp": datetime.now(UTC).isoformat(),
        "combined": round(combined.percent, 2),
        "domains": {
            domain.label: round(summaries[domain.label].percent, 2)
            for domain in CORE_DOMAINS
            if summaries[domain.label].total_lines > 0
        },
    }
    history = _read_coverage_history()
    history.append(record)
    _save_coverage_snapshot(history)


def _compute_climbing_status(history: list[dict]) -> dict:
    global CURRENT_THRESHOLD, NEXT_THRESHOLD
    if CURRENT_THRESHOLD is None:
        repo_root = Path(__file__).resolve().parents[1]
        CURRENT_THRESHOLD = _read_pyproject_fail_under(repo_root)
    if NEXT_THRESHOLD is None:
        NEXT_THRESHOLD = CURRENT_THRESHOLD

    recent = history[-CLIMBING_N:] if len(history) >= CLIMBING_N else history
    recent_combined = [r["combined"] for r in recent if "combined" in r]

    if NEXT_THRESHOLD == CURRENT_THRESHOLD or len(recent_combined) < CLIMBING_N:
        return {
            "current_threshold": CURRENT_THRESHOLD,
            "next_threshold": None,
            "upgrade_triggered": "pending",
            "reason": "next_threshold not yet measured (step 1 pending) or insufficient CI history",
            "recent_combined_values": recent_combined,
            "n": CLIMBING_N,
            "m": CLIMBING_M,
            "buffer_pp": CLIMBING_BUFFER_PP,
        }

    target = NEXT_THRESHOLD + CLIMBING_BUFFER_PP
    all_above = all(v >= target for v in recent_combined)

    return {
        "current_threshold": CURRENT_THRESHOLD,
        "next_threshold": NEXT_THRESHOLD,
        "upgrade_triggered": all_above,
        "target_for_upgrade": round(target, 2),
        "recent_combined_values": recent_combined,
        "n": CLIMBING_N,
        "m": CLIMBING_M,
        "buffer_pp": CLIMBING_BUFFER_PP,
    }


def _emit_climbing_check(status: dict) -> None:
    print("<climbing_check>")
    print(f"current_threshold={status['current_threshold']}")
    print(f"next_threshold={status['next_threshold']}")
    print(f"upgrade_triggered={status['upgrade_triggered']}")
    if "target_for_upgrade" in status:
        print(f"target_for_upgrade={status['target_for_upgrade']}")
    if status.get("reason"):
        print(f"reason={status['reason']}")
    recent = status.get("recent_combined_values", [])
    if recent:
        print(f"recent_combined_values={'|'.join(f'{v:.2f}' for v in recent)}")
    print(f"n={status['n']} m={status['m']} buffer_pp={status['buffer_pp']}")
    print("</climbing_check>")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("coverage_xml", nargs="?", default="coverage.xml")
    parser.add_argument(
        "--soft-gate",
        action="store_true",
        help="Apply the built-in report-only thresholds for core domains and combined coverage",
    )
    parser.add_argument(
        "--report-under",
        action="append",
        default=[],
        metavar="DOMAIN=PERCENT",
        help="Report-only threshold; prints warnings but does not fail the command",
    )
    parser.add_argument(
        "--fail-under",
        action="append",
        default=[],
        metavar="DOMAIN=PERCENT",
        help="Optional threshold that fails the command, e.g. converter/=55 or combined=57",
    )
    parser.add_argument(
        "--save-snapshot",
        action="store_true",
        help="Append current coverage data to the history file for climbing curve tracking",
    )
    args = parser.parse_args(argv)

    coverage_xml = Path(args.coverage_xml)
    if not coverage_xml.is_file():
        print(f"coverage XML not found: {coverage_xml}", file=sys.stderr)
        return 1

    try:
        report_thresholds = _parse_fail_under(args.report_under)
        thresholds = _parse_fail_under(args.fail_under)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if args.soft_gate:
        report_thresholds = dict(SOFT_GATE_THRESHOLDS) | report_thresholds

    summaries = _summaries_from_coverage_xml(coverage_xml)
    combined_summary = _print_report(summaries)

    exit_code = 0
    if combined_summary.total_lines > 0:
        print(
            f"combined core domains: {combined_summary.percent:.2f}% "
            f"({combined_summary.covered_lines}/{combined_summary.total_lines} lines)"
        )

    exit_code |= _evaluate_thresholds(
        thresholds=report_thresholds,
        prefix="report-threshold",
        annotation_level="warning",
        fail_on_breach=False,
        summaries=summaries,
        combined_summary=combined_summary,
    )
    exit_code |= _evaluate_thresholds(
        thresholds=thresholds,
        prefix="threshold-failed",
        annotation_level="error",
        fail_on_breach=True,
        summaries=summaries,
        combined_summary=combined_summary,
    )

    if any(not summaries[domain.label].files for domain in CORE_DOMAINS):
        _emit_annotation("error", "core-domain-coverage missing one or more required domains in coverage XML")
        exit_code = 1

    if args.save_snapshot:
        _append_coverage_snapshot(summaries, combined_summary)
        print(f"[snapshot] saved to {_HISTORY_PATH}")

    if args.soft_gate:
        history = _read_coverage_history()
        status = _compute_climbing_status(history)
        _emit_climbing_check(status)

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
