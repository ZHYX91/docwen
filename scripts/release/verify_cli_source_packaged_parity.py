"""Compare protocol 3 discovery and content-inspection fixtures across CLI lanes."""

from __future__ import annotations

import argparse
import copy
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl

_TEMPLATE_ID_PATTERN = re.compile(r"^template\.(?:docx|xlsx)\.[0-9a-f]{64}$")
_TEMPLATE_RESOURCE_FIELDS = frozenset({"id", "name", "target", "description", "path", "size_bytes", "modified_ns"})


@dataclass(frozen=True, slots=True)
class Fixture:
    name: str
    argv: tuple[str, ...]
    expected_exit: int
    expected_data: tuple[tuple[str, object], ...] = ()


def _write_xlsx(path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    if worksheet is None:
        raise RuntimeError("cli_parity_workbook_has_no_active_sheet")
    worksheet["A1"] = "name"
    worksheet["B1"] = "value"
    worksheet["A2"] = "alpha"
    worksheet["B2"] = 1
    workbook.save(path)
    workbook.close()


def _build_fixtures(root: Path) -> tuple[Fixture, ...]:
    plain_text = "Plain text without Markdown markers.\nSecond plain line.\n"
    markdown_text = "# Heading\n\n**bold** text\n"
    plain_txt = root / "纯文本内容.txt"
    plain_md = root / "纯文本内容.md"
    markdown_txt = root / "Markdown 内容.txt"
    markdown_md = root / "Markdown 内容.md"
    plain_txt.write_text(plain_text, encoding="utf-8")
    plain_md.write_text(plain_text, encoding="utf-8")
    markdown_txt.write_text(markdown_text, encoding="utf-8")
    markdown_md.write_text(markdown_text, encoding="utf-8")

    disguised_xlsx = root / "实际为 XLSX 的文本后缀.txt"
    _write_xlsx(disguised_xlsx)
    ordinary_zip = root / "普通 ZIP 伪装文档.docx"
    with zipfile.ZipFile(ordinary_zip, "w") as package:
        package.writestr("hello.txt", "hello")
    corrupt_ooxml = root / "损坏 OOXML 伪装表格.xlsx"
    corrupt_ooxml.write_bytes(b"PK\x03\x04not-a-valid-central-directory")
    missing = root / "不存在 文件.md"

    return (
        Fixture("info", ("info", "--json"), 0),
        Fixture("schema-convert", ("schema", "convert", "--json"), 0),
        Fixture("schema-number-markdown", ("schema", "number", "markdown", "--json"), 0),
        Fixture("resources-formats", ("resources", "list", "formats", "--json", "--quiet"), 0),
        Fixture(
            "resources-optimizations",
            ("resources", "list", "optimizations", "--json"),
            0,
            (
                ("resource", "optimizations"),
                ("contract", {"id": "docwen.optimizations", "version": 1}),
            ),
        ),
        Fixture(
            "resources-templates",
            ("resources", "list", "templates", "--json", "--quiet"),
            0,
            (("type", "templates"),),
        ),
        Fixture(
            "resources-numbering-schemes",
            ("resources", "list", "numbering-schemes", "--json", "--quiet"),
            0,
        ),
        Fixture(
            "inspect-plain-text-as-txt",
            ("inspect", str(plain_txt), "--json", "--quiet"),
            0,
            (
                ("declared_format", "txt"),
                ("detected_format", "txt"),
                ("relation", "exact_match"),
                ("decision", "allow"),
            ),
        ),
        Fixture(
            "inspect-plain-text-as-markdown",
            ("inspect", str(plain_md), "--json", "--quiet"),
            0,
            (
                ("declared_format", "markdown"),
                ("detected_format", "txt"),
                ("relation", "compatible_text"),
                ("decision", "allow_with_warning"),
            ),
        ),
        Fixture(
            "inspect-markdown-as-txt",
            ("inspect", str(markdown_txt), "--json", "--quiet"),
            0,
            (
                ("declared_format", "txt"),
                ("detected_format", "markdown"),
                ("relation", "compatible_text"),
                ("decision", "allow_with_warning"),
            ),
        ),
        Fixture(
            "inspect-markdown-as-markdown",
            ("inspect", str(markdown_md), "--json", "--quiet"),
            0,
            (
                ("declared_format", "markdown"),
                ("detected_format", "markdown"),
                ("relation", "equivalent_alias"),
                ("decision", "allow"),
            ),
        ),
        Fixture(
            "inspect-xlsx-as-txt",
            ("inspect", str(disguised_xlsx), "--json", "--quiet"),
            0,
            (
                ("declared_format", "txt"),
                ("detected_format", "xlsx"),
                ("relation", "cross_family_mismatch"),
                ("decision", "require_explicit_acceptance"),
            ),
        ),
        Fixture(
            "inspect-ordinary-zip-as-docx",
            ("inspect", str(ordinary_zip), "--json", "--quiet"),
            0,
            (
                ("declared_format", "docx"),
                ("detected_format", "zip"),
                ("decision", "block"),
                ("reason_code", "FILE_CONTAINER_INVALID"),
            ),
        ),
        Fixture(
            "inspect-corrupt-ooxml-as-xlsx",
            ("inspect", str(corrupt_ooxml), "--json", "--quiet"),
            0,
            (
                ("declared_format", "xlsx"),
                ("decision", "block"),
                ("reason_code", "FILE_CONTAINER_INVALID"),
            ),
        ),
        Fixture("inspect-missing", ("inspect", str(missing), "--json", "--quiet"), 2),
    )


def _run(command: list[str], *, cwd: Path, env: dict[str, str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        cwd=cwd,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="strict",
        check=False,
    )


def _load_payload(result: subprocess.CompletedProcess[str], fixture: Fixture, lane: str) -> dict[str, Any]:
    if result.returncode != fixture.expected_exit:
        raise RuntimeError(
            f"{fixture.name}/{lane}: expected exit {fixture.expected_exit}, got {result.returncode}\n"
            f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )
    if result.stderr:
        raise RuntimeError(f"{fixture.name}/{lane}: machine mode wrote stderr:\n{result.stderr}")
    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"{fixture.name}/{lane}: invalid JSON:\n{result.stdout}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError(f"{fixture.name}/{lane}: top-level JSON is not an object")
    if fixture.expected_data:
        data = payload.get("data")
        if not isinstance(data, dict):
            raise RuntimeError(f"{fixture.name}/{lane}: expected a data object, got {payload}")
        for field, expected_value in fixture.expected_data:
            if data.get(field) != expected_value:
                raise RuntimeError(
                    f"{fixture.name}/{lane}: expected data.{field}={expected_value!r}, got {data.get(field)!r}"
                )
    return payload


def _template_resources_by_id(payload: dict[str, Any], *, lane: str) -> dict[str, dict[str, Any]]:
    if payload.get("success") is not True or payload.get("command") != "resources list":
        raise RuntimeError(f"resources-templates/{lane}: invalid success envelope: {payload}")
    data = payload.get("data")
    if not isinstance(data, dict) or set(data) != {"type", "resources", "total"}:
        raise RuntimeError(f"resources-templates/{lane}: invalid data contract: {data}")
    if data.get("type") != "templates":
        raise RuntimeError(f"resources-templates/{lane}: expected type='templates'")
    resources = data.get("resources")
    total = data.get("total")
    if not isinstance(resources, list) or not resources:
        raise RuntimeError(f"resources-templates/{lane}: no template resources were returned")
    if type(total) is not int or total != len(resources):
        raise RuntimeError(f"resources-templates/{lane}: total does not match resources")

    by_id: dict[str, dict[str, Any]] = {}
    for resource in resources:
        if not isinstance(resource, dict) or set(resource) != _TEMPLATE_RESOURCE_FIELDS:
            raise RuntimeError(f"resources-templates/{lane}: invalid resource contract: {resource}")
        template_id = resource.get("id")
        target = resource.get("target")
        if (
            not isinstance(template_id, str)
            or _TEMPLATE_ID_PATTERN.fullmatch(template_id) is None
            or target not in {"docx", "xlsx"}
            or not template_id.startswith(f"template.{target}.")
        ):
            raise RuntimeError(f"resources-templates/{lane}: invalid canonical template ID: {resource}")
        if template_id in by_id:
            raise RuntimeError(f"resources-templates/{lane}: duplicate canonical template ID: {template_id}")
        by_id[template_id] = resource
    return by_id


def _select_canonical_template_id(payload: dict[str, Any], *, lane: str) -> str:
    """Select a stable token using only IDs returned by resource discovery."""

    return min(_template_resources_by_id(payload, lane=lane))


def _normalize_template_storage_metadata(payload: dict[str, Any], *, lane: str) -> dict[str, Any]:
    """Normalize only lane-local template path and filesystem timestamp metadata."""

    data = payload.get("data")
    if not isinstance(data, dict) or data.get("type") != "templates":
        return payload
    if payload.get("command") == "resources list":
        raw_resources = data.get("resources")
    elif payload.get("command") == "resources show":
        raw_resources = [data.get("resource")]
    else:
        return payload
    if not isinstance(raw_resources, list):
        raise RuntimeError(f"templates/{lane}: resource collection is not a list")

    normalized = copy.deepcopy(payload)
    normalized_data = normalized["data"]
    assert isinstance(normalized_data, dict)
    normalized_resources = (
        normalized_data.get("resources")
        if payload.get("command") == "resources list"
        else [normalized_data.get("resource")]
    )
    assert isinstance(normalized_resources, list)
    for resource in normalized_resources:
        if not isinstance(resource, dict):
            raise RuntimeError(f"templates/{lane}: resource is not an object")
        resource_path = resource.get("path")
        modified_ns = resource.get("modified_ns")
        if not isinstance(resource_path, str) or not resource_path or type(modified_ns) is not int:
            raise RuntimeError(f"templates/{lane}: invalid lane-local storage metadata: {resource}")
        resource["path"] = "<verified-lane-template-path>"
        resource["modified_ns"] = "<verified-lane-template-mtime>"
    return normalized


def _normalize_verified_lane_bootstrap(payload: dict[str, Any], *, lane: str) -> dict[str, Any]:
    """Normalize only the already-verified source/frozen guard owner."""

    data = payload.get("data")
    security = data.get("security") if isinstance(data, dict) else None
    guard = security.get("dependency_egress_guard") if isinstance(security, dict) else None
    if guard is None:
        return payload
    if not isinstance(guard, dict):
        raise RuntimeError(f"{lane}: dependency egress guard is not an object")

    expected_bootstrap = {
        "source": "composition_root",
        "packaged": "pyinstaller_runtime_hook",
    }.get(lane)
    if expected_bootstrap is None:
        raise ValueError(f"unknown parity lane: {lane}")
    if guard.get("bootstrap") != expected_bootstrap:
        raise RuntimeError(
            f"{lane}: expected dependency egress bootstrap {expected_bootstrap!r}, got {guard.get('bootstrap')!r}"
        )

    normalized = copy.deepcopy(payload)
    normalized_data = normalized["data"]
    assert isinstance(normalized_data, dict)
    normalized_security = normalized_data["security"]
    assert isinstance(normalized_security, dict)
    normalized_guard = normalized_security["dependency_egress_guard"]
    assert isinstance(normalized_guard, dict)
    normalized_guard["bootstrap"] = "<verified-lane-bootstrap>"
    return normalized


def _normalize_for_parity(payload: dict[str, Any], *, lane: str) -> dict[str, Any]:
    normalized = _normalize_verified_lane_bootstrap(payload, lane=lane)
    return _normalize_template_storage_metadata(normalized, lane=lane)


def _verify_template_show_parity(
    template_id: str,
    *,
    source_list_payload: dict[str, Any],
    packaged_list_payload: dict[str, Any],
    source_prefix: list[str],
    packaged_prefix: list[str],
    cwd: Path,
    env: dict[str, str],
) -> None:
    source_listed = _template_resources_by_id(source_list_payload, lane="source")
    packaged_listed = _template_resources_by_id(packaged_list_payload, lane="packaged")
    if template_id not in source_listed or template_id not in packaged_listed:
        raise RuntimeError(f"resources-template-show: listed canonical ID is missing: {template_id}")

    fixture = Fixture(
        "resources-template-show",
        ("resources", "show", "templates", template_id, "--json", "--quiet"),
        0,
        (("type", "templates"),),
    )
    source = _run([*source_prefix, *fixture.argv], cwd=cwd, env=env)
    packaged = _run([*packaged_prefix, *fixture.argv], cwd=cwd, env=env)
    source_payload = _load_payload(source, fixture, "source")
    packaged_payload = _load_payload(packaged, fixture, "packaged")

    for lane, payload, listed in (
        ("source", source_payload, source_listed[template_id]),
        ("packaged", packaged_payload, packaged_listed[template_id]),
    ):
        if payload.get("success") is not True or payload.get("command") != "resources show":
            raise RuntimeError(f"{fixture.name}/{lane}: invalid success envelope: {payload}")
        data = payload.get("data")
        if not isinstance(data, dict) or set(data) != {"type", "resource"}:
            raise RuntimeError(f"{fixture.name}/{lane}: invalid data contract: {data}")
        resource = data.get("resource")
        if not isinstance(resource, dict) or set(resource) != _TEMPLATE_RESOURCE_FIELDS:
            raise RuntimeError(f"{fixture.name}/{lane}: invalid resource contract: {resource}")
        if resource != listed:
            raise RuntimeError(f"{fixture.name}/{lane}: shown resource does not exactly match its listed resource")

    normalized_source = _normalize_for_parity(source_payload, lane="source")
    normalized_packaged = _normalize_for_parity(packaged_payload, lane="packaged")
    if normalized_source != normalized_packaged:
        raise RuntimeError(
            f"{fixture.name}: source and packaged payloads differ\n"
            f"source={json.dumps(source_payload, ensure_ascii=False, sort_keys=True)}\n"
            f"packaged={json.dumps(packaged_payload, ensure_ascii=False, sort_keys=True)}"
        )


def verify(binary_dir: Path) -> None:
    binary_name = "DocWenCLI.exe" if os.name == "nt" else "DocWenCLI"
    binary = (binary_dir / binary_name).resolve()
    if not binary.is_file():
        raise FileNotFoundError(f"Packaged CLI not found: {binary}")

    with tempfile.TemporaryDirectory(prefix="docwen-cli-parity-") as raw_temp:
        root = Path(raw_temp) / "资料 空格"
        root.mkdir()
        fixtures = _build_fixtures(root)

        env = os.environ.copy()
        env.update(
            {
                "PYTHONUTF8": "1",
                "PYTHONIOENCODING": "utf-8",
                "DOCWEN_CONFIG_DIR": str(root / "config"),
                "DOCWEN_LOG_DIR": str(root / "logs"),
                "DOCWEN_LOG_TO_TEMP": "",
            }
        )
        source_prefix = [sys.executable, "-m", "docwen_bundle.cli_entry"]
        packaged_prefix = [str(binary)]
        template_list_payloads: tuple[dict[str, Any], dict[str, Any]] | None = None

        for fixture in fixtures:
            source = _run([*source_prefix, *fixture.argv], cwd=root, env=env)
            packaged = _run([*packaged_prefix, *fixture.argv], cwd=root, env=env)
            source_payload = _load_payload(source, fixture, "source")
            packaged_payload = _load_payload(packaged, fixture, "packaged")
            normalized_source = _normalize_for_parity(source_payload, lane="source")
            normalized_packaged = _normalize_for_parity(packaged_payload, lane="packaged")
            if normalized_source != normalized_packaged:
                raise RuntimeError(
                    f"{fixture.name}: source and packaged payloads differ\n"
                    f"source={json.dumps(source_payload, ensure_ascii=False, sort_keys=True)}\n"
                    f"packaged={json.dumps(packaged_payload, ensure_ascii=False, sort_keys=True)}"
                )
            if fixture.name == "resources-templates":
                template_list_payloads = (source_payload, packaged_payload)

        if template_list_payloads is None:
            raise RuntimeError("resources-templates fixture did not run")
        source_template_list, packaged_template_list = template_list_payloads
        source_template_id = _select_canonical_template_id(source_template_list, lane="source")
        packaged_template_id = _select_canonical_template_id(packaged_template_list, lane="packaged")
        if source_template_id != packaged_template_id:
            raise RuntimeError("resources-template-show: source and packaged list selected different canonical IDs")
        _verify_template_show_parity(
            source_template_id,
            source_list_payload=source_template_list,
            packaged_list_payload=packaged_template_list,
            source_prefix=source_prefix,
            packaged_prefix=packaged_prefix,
            cwd=root,
            env=env,
        )

    print(f"cli_source_packaged_parity_ok: {len(fixtures) + 1} fixtures -> {binary.name}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary-dir", type=Path, required=True)
    args = parser.parse_args()
    verify(args.binary_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
