from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree

GUI_PATH_PREFIX = "packages/apps/gui/src/docwen_gui/"


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


def _normalize_coverage_filename(filename: str) -> str:
    normalized = filename.replace("\\", "/").strip()
    while normalized.startswith("./"):
        normalized = normalized[2:]
    return normalized


def _is_gui_filename(filename: str) -> bool:
    normalized = _normalize_coverage_filename(filename)
    return normalized.startswith(GUI_PATH_PREFIX) or f"/{GUI_PATH_PREFIX}" in normalized


def _summary_from_coverage_xml(coverage_xml: Path) -> CoverageSummary:
    root = ElementTree.parse(coverage_xml).getroot()
    summary = CoverageSummary(files=set())

    for class_node in root.findall(".//class"):
        filename = class_node.attrib.get("filename", "")
        if not filename or not _is_gui_filename(filename):
            continue

        lines = class_node.findall("./lines/line")
        total_lines = len(lines)
        covered_lines = sum(1 for line in lines if int(line.attrib.get("hits", "0")) > 0)
        summary.files.add(_normalize_coverage_filename(filename))
        summary.covered_lines += covered_lines
        summary.total_lines += total_lines

    return summary


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("coverage_xml", nargs="?", default="coverage-gui.xml")
    parser.add_argument(
        "--fail-under",
        type=float,
        default=None,
        help="Optional GUI coverage threshold for future use.",
    )
    args = parser.parse_args(argv)

    coverage_xml = Path(args.coverage_xml)
    if not coverage_xml.is_file():
        print(f"coverage XML not found: {coverage_xml}", file=sys.stderr)
        return 1

    summary = _summary_from_coverage_xml(coverage_xml)
    print("==> gui-coverage")
    if not summary.files:
        print("[missing] gui/: no matching files found in coverage XML")
        return 1

    print(
        f"gui/: {summary.percent:.2f}% "
        f"({summary.covered_lines}/{summary.total_lines} lines, {len(summary.files)} files)"
    )

    if args.fail_under is not None and summary.percent < args.fail_under:
        print(
            f"[threshold-failed] gui/: {summary.percent:.2f}% < required {args.fail_under:.2f}%",
            file=sys.stderr,
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
