from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract


def _success_payload(command: str, data: dict[str, object]) -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": True,
        "command": command,
        "data": data,
        "error": None,
    }


def _blocked_payload(input_path: Path) -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": False,
        "command": "convert",
        "data": None,
        "error": {
            "category": "invalid_input",
            "code": "file_container_invalid",
            "message": "invalid container",
            "details": {
                "file": str(input_path),
                "admission": {
                    "decision": "block",
                    "reason_code": "FILE_CONTAINER_INVALID",
                },
            },
            "hint": None,
        },
    }


def test_packaged_cli_content_first_smoke_covers_success_and_hard_blocks(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    binary_path = tmp_path / "DocWenCLI.exe"
    binary_path.write_text("placeholder", encoding="utf-8")
    calls: list[tuple[str, ...]] = []
    inspections = {
        "实际为 XLSX 的文本后缀.txt": {
            "declared_format": "txt",
            "detected_format": "xlsx",
            "relation": "cross_family_mismatch",
            "decision": "require_explicit_acceptance",
        },
        "纯文本内容.txt": {
            "declared_format": "txt",
            "detected_format": "txt",
            "relation": "exact_match",
            "decision": "allow",
        },
        "纯文本内容.md": {
            "declared_format": "markdown",
            "detected_format": "txt",
            "relation": "compatible_text",
            "decision": "allow_with_warning",
        },
        "Markdown 内容.txt": {
            "declared_format": "txt",
            "detected_format": "markdown",
            "relation": "compatible_text",
            "decision": "allow_with_warning",
        },
        "Markdown 内容.md": {
            "declared_format": "markdown",
            "detected_format": "markdown",
            "relation": "equivalent_alias",
            "decision": "allow",
        },
    }
    blocked_names = {"普通 ZIP 伪装文档.docx", "损坏 OOXML 伪装表格.xlsx"}

    def fake_write_xlsx(path: Path) -> None:
        path.write_bytes(b"xlsx")

    def fake_run(binary: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        if args[0] == "inspect":
            payload = _success_payload("inspect", inspections[Path(args[1]).name])
            return subprocess.CompletedProcess([str(binary), *args], 0, stdout=json.dumps(payload), stderr="")

        assert args[0] == "convert"
        input_path = Path(args[1])
        output_path = Path(args[args.index("--output") + 1])
        if input_path.name in blocked_names:
            payload = _blocked_payload(input_path)
            return subprocess.CompletedProcess([str(binary), *args], 2, stdout=json.dumps(payload), stderr="")
        output_path.write_text("| name | value |\n| --- | --- |\n| alpha | 1 |\n", encoding="utf-8")
        payload = _success_payload("convert", {"output": str(output_path)})
        return subprocess.CompletedProcess([str(binary), *args], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_cli, "_write_xlsx", fake_write_xlsx)
    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    output = verify_packaged_cli._run_content_first_contract_smoke(binary_path, work_dir=tmp_path)

    assert output.name == "伪装表格转换结果.md"
    assert output.is_file()
    inspect_calls = [call for call in calls if call[0] == "inspect"]
    convert_calls = [call for call in calls if call[0] == "convert"]
    assert {Path(call[1]).name for call in inspect_calls} == set(inspections)
    assert {Path(call[1]).name for call in convert_calls} == {
        "实际为 XLSX 的文本后缀.txt",
        *blocked_names,
    }
    assert all("--use-detected-format" in call for call in convert_calls)
    assert not (tmp_path / "普通 ZIP 不应生成.md").exists()
    assert not (tmp_path / "损坏 OOXML 不应生成.md").exists()


def test_packaged_cli_failure_loader_rejects_false_success() -> None:
    from scripts.release import verify_packaged_cli

    payload = _success_payload("convert", {"output": "should-not-exist.md"})
    proc = subprocess.CompletedProcess(
        ["DocWenCLI.exe", "convert"],
        0,
        stdout=json.dumps(payload),
        stderr="",
    )

    with pytest.raises(RuntimeError, match="unexpectedly succeeded"):
        verify_packaged_cli._load_json_failure_payload(proc, command_name="blocked convert")


def test_source_packaged_parity_fixtures_cover_content_first_inspection(tmp_path: Path) -> None:
    from scripts.release import verify_cli_source_packaged_parity

    from docwen_core.detection import inspect_file

    fixtures = verify_cli_source_packaged_parity._build_fixtures(tmp_path)
    fixture_names = {fixture.name for fixture in fixtures}
    assert {
        "resources-templates",
        "inspect-plain-text-as-txt",
        "inspect-plain-text-as-markdown",
        "inspect-markdown-as-txt",
        "inspect-markdown-as-markdown",
        "inspect-xlsx-as-txt",
        "inspect-ordinary-zip-as-docx",
        "inspect-corrupt-ooxml-as-xlsx",
    } <= fixture_names
    templates_fixture = next(fixture for fixture in fixtures if fixture.name == "resources-templates")
    assert templates_fixture.argv == ("resources", "list", "templates", "--json", "--quiet")

    for fixture in fixtures:
        if fixture.argv[0] != "inspect" or not fixture.expected_data:
            continue
        input_path = Path(fixture.argv[1])
        data = inspect_file(str(input_path)).to_dict()
        result = subprocess.CompletedProcess(
            ["DocWenCLI.exe", *fixture.argv],
            fixture.expected_exit,
            stdout=json.dumps(_success_payload("inspect", data)),
            stderr="",
        )
        verify_cli_source_packaged_parity._load_payload(result, fixture, "source")


def test_source_packaged_parity_rejects_matching_but_wrong_inspection_semantics() -> None:
    from scripts.release import verify_cli_source_packaged_parity

    fixture = verify_cli_source_packaged_parity.Fixture(
        "inspect-blocked",
        ("inspect", "blocked.docx", "--json", "--quiet"),
        0,
        (("decision", "block"),),
    )
    result = subprocess.CompletedProcess(
        ["DocWenCLI.exe", *fixture.argv],
        0,
        stdout=json.dumps(_success_payload("inspect", {"decision": "allow"})),
        stderr="",
    )

    with pytest.raises(RuntimeError, match=r"expected data\.decision='block'"):
        verify_cli_source_packaged_parity._load_payload(result, fixture, "packaged")


def test_source_packaged_parity_normalizes_only_verified_egress_bootstrap() -> None:
    from scripts.release import verify_cli_source_packaged_parity

    source = _success_payload(
        "resources list",
        {
            "security": {
                "dependency_egress_guard": {
                    "state": "enforced",
                    "bootstrap": "composition_root",
                    "policy": "deny_dns_and_ip",
                }
            }
        },
    )
    packaged = _success_payload(
        "resources list",
        {
            "security": {
                "dependency_egress_guard": {
                    "state": "enforced",
                    "bootstrap": "pyinstaller_runtime_hook",
                    "policy": "deny_dns_and_ip",
                }
            }
        },
    )

    normalized_source = verify_cli_source_packaged_parity._normalize_verified_lane_bootstrap(
        source,
        lane="source",
    )
    normalized_packaged = verify_cli_source_packaged_parity._normalize_verified_lane_bootstrap(
        packaged,
        lane="packaged",
    )

    assert normalized_source == normalized_packaged
    assert source["data"]["security"]["dependency_egress_guard"]["bootstrap"] == "composition_root"
    assert packaged["data"]["security"]["dependency_egress_guard"]["bootstrap"] == "pyinstaller_runtime_hook"


def test_source_packaged_parity_rejects_unexpected_egress_bootstrap() -> None:
    from scripts.release import verify_cli_source_packaged_parity

    payload = _success_payload(
        "resources list",
        {
            "security": {
                "dependency_egress_guard": {
                    "state": "enforced",
                    "bootstrap": "composition_root",
                }
            }
        },
    )

    with pytest.raises(RuntimeError, match="expected dependency egress bootstrap 'pyinstaller_runtime_hook'"):
        verify_cli_source_packaged_parity._normalize_verified_lane_bootstrap(payload, lane="packaged")


def test_source_packaged_template_show_uses_one_discovered_id_without_name_inference(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_cli_source_packaged_parity

    template_id = f"template.docx.{'2' * 64}"

    def resource(*, lane: str) -> dict[str, object]:
        return {
            "id": template_id,
            "name": "A display name that must not be used as a selector",
            "target": "docx",
            "description": "Template description",
            "path": str(tmp_path / lane / "renamed-template.docx"),
            "size_bytes": 123,
            "modified_ns": 100 if lane == "source" else 200,
        }

    source_resource = resource(lane="source")
    packaged_resource = resource(lane="packaged")
    source_list = _success_payload(
        "resources list",
        {"type": "templates", "resources": [source_resource], "total": 1},
    )
    packaged_list = _success_payload(
        "resources list",
        {"type": "templates", "resources": [packaged_resource], "total": 1},
    )
    calls: list[tuple[str, ...]] = []

    def fake_run(command: list[str], *, cwd: Path, env: dict[str, str]) -> subprocess.CompletedProcess[str]:
        del cwd, env
        calls.append(tuple(command))
        lane = command[0]
        shown = source_resource if lane == "source" else packaged_resource
        return subprocess.CompletedProcess(
            command,
            0,
            stdout=json.dumps(_success_payload("resources show", {"type": "templates", "resource": shown})),
            stderr="",
        )

    monkeypatch.setattr(verify_cli_source_packaged_parity, "_run", fake_run)

    selected = verify_cli_source_packaged_parity._select_canonical_template_id(
        source_list,
        lane="source",
    )
    assert selected == template_id
    verify_cli_source_packaged_parity._verify_template_show_parity(
        selected,
        source_list_payload=source_list,
        packaged_list_payload=packaged_list,
        source_prefix=["source"],
        packaged_prefix=["packaged"],
        cwd=tmp_path,
        env={},
    )

    expected_suffix = ("resources", "show", "templates", template_id, "--json", "--quiet")
    assert calls == [
        ("source", *expected_suffix),
        ("packaged", *expected_suffix),
    ]


def test_source_packaged_template_parity_normalizes_only_lane_local_storage_metadata() -> None:
    from scripts.release import verify_cli_source_packaged_parity

    template_id = f"template.xlsx.{'3' * 64}"
    source = _success_payload(
        "resources show",
        {
            "type": "templates",
            "resource": {
                "id": template_id,
                "name": "Stable",
                "target": "xlsx",
                "description": "Stable Excel template",
                "path": "D:/source/templates/Stable.xlsx",
                "size_bytes": 99,
                "modified_ns": 111,
            },
        },
    )
    packaged = json.loads(json.dumps(source))
    packaged["data"]["resource"]["path"] = "D:/package/templates/Stable.xlsx"
    packaged["data"]["resource"]["modified_ns"] = 222

    assert verify_cli_source_packaged_parity._normalize_template_storage_metadata(
        source,
        lane="source",
    ) == verify_cli_source_packaged_parity._normalize_template_storage_metadata(
        packaged,
        lane="packaged",
    )
    assert source["data"]["resource"]["path"] == "D:/source/templates/Stable.xlsx"


def test_source_packaged_template_show_rejects_drift_from_its_list_item(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_cli_source_packaged_parity

    template_id = f"template.docx.{'4' * 64}"
    listed = {
        "id": template_id,
        "name": "Stable",
        "target": "docx",
        "description": "Stable template",
        "path": str(tmp_path / "Stable.docx"),
        "size_bytes": 10,
        "modified_ns": 20,
    }
    list_payload = _success_payload(
        "resources list",
        {"type": "templates", "resources": [listed], "total": 1},
    )
    drifted = dict(listed)
    drifted["description"] = "Drifted template"

    monkeypatch.setattr(
        verify_cli_source_packaged_parity,
        "_run",
        lambda command, cwd, env: subprocess.CompletedProcess(
            command,
            0,
            stdout=json.dumps(
                _success_payload(
                    "resources show",
                    {
                        "type": "templates",
                        "resource": drifted if command[0] == "source" else listed,
                    },
                )
            ),
            stderr="",
        ),
    )

    with pytest.raises(RuntimeError, match="shown resource does not exactly match"):
        verify_cli_source_packaged_parity._verify_template_show_parity(
            template_id,
            source_list_payload=list_payload,
            packaged_list_payload=list_payload,
            source_prefix=["source"],
            packaged_prefix=["packaged"],
            cwd=tmp_path,
            env={},
        )
