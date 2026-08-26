from __future__ import annotations

import argparse
import sys
import tomllib
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree


def _discover_source_packages(repo_root: Path) -> list[str]:
    packages_root = repo_root / "packages"
    discovered: set[str] = set()
    for package_config in packages_root.rglob("pyproject.toml"):
        source_root = package_config.parent / "src"
        if not source_root.is_dir():
            continue
        for candidate in source_root.iterdir():
            if candidate.is_dir() and (candidate / "__init__.py").is_file():
                discovered.add(candidate.name)
    return sorted(discovered)


def _configured_source_packages(repo_root: Path) -> list[str]:
    pyproject_path = repo_root / "pyproject.toml"
    data = tomllib.loads(pyproject_path.read_text(encoding="utf-8"))
    raw_sources = data["tool"]["coverage"]["run"]["source"]
    if not isinstance(raw_sources, list) or not all(isinstance(source, str) for source in raw_sources):
        raise ValueError("coverage_source_manifest_invalid: tool.coverage.run.source must be a string list")
    if len(raw_sources) != len(set(raw_sources)):
        raise ValueError("coverage_source_manifest_invalid: duplicate package names")
    invalid = sorted(source for source in raw_sources if not source.isidentifier())
    if invalid:
        raise ValueError("coverage_source_manifest_invalid: expected import package names: " + ", ".join(invalid))
    return sorted(raw_sources)


def _reported_source_packages(coverage_xml: Path, configured: list[str]) -> set[str]:
    root = ElementTree.parse(coverage_xml).getroot()
    filenames = [node.attrib.get("filename", "") for node in root.findall(".//class")]
    if not any(filenames):
        raise ValueError("coverage_no_data: coverage XML contains no reported source files")

    reported: set[str] = set()
    configured_set = set(configured)
    for filename in filenames:
        parts = set(PurePosixPath(filename.replace("\\", "/")).parts)
        reported.update(parts & configured_set)
    return reported


def check_coverage_source_manifest(repo_root: Path, coverage_xml: Path) -> list[str]:
    errors: list[str] = []
    try:
        configured = _configured_source_packages(repo_root)
    except (KeyError, TypeError, ValueError, tomllib.TOMLDecodeError) as exc:
        return [str(exc)]

    discovered = _discover_source_packages(repo_root)
    missing_from_config = sorted(set(discovered) - set(configured))
    unknown_in_config = sorted(set(configured) - set(discovered))
    if missing_from_config:
        errors.append("coverage_source_manifest_missing: " + ", ".join(missing_from_config))
    if unknown_in_config:
        errors.append("coverage_source_manifest_unknown: " + ", ".join(unknown_in_config))
    if errors:
        return errors

    if not coverage_xml.is_file():
        return [f"coverage_xml_missing: {coverage_xml}"]
    try:
        reported = _reported_source_packages(coverage_xml, configured)
    except (ElementTree.ParseError, OSError, ValueError) as exc:
        return [str(exc)]

    missing_from_report = sorted(set(configured) - reported)
    if missing_from_report:
        errors.append("coverage_module_not_reported: " + ", ".join(missing_from_report))
    return errors


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("coverage_xml", nargs="?", default="coverage.xml", type=Path)
    parser.add_argument("--repo-root", default=Path(__file__).resolve().parents[1], type=Path)
    args = parser.parse_args(argv)

    repo_root = args.repo_root.resolve()
    coverage_xml = args.coverage_xml
    if not coverage_xml.is_absolute():
        coverage_xml = (repo_root / coverage_xml).resolve()

    errors = check_coverage_source_manifest(repo_root, coverage_xml)
    if errors:
        for error in errors:
            print(error, file=sys.stderr)
        return 1

    configured = _configured_source_packages(repo_root)
    print(f"ok: coverage source manifest reports all {len(configured)} configured packages")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
