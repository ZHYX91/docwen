"""Focused packaged CLI template list/show/convert contracts."""

from __future__ import annotations

import json
import subprocess
import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract


def _template_resource(*, resource_id: str, name: str, target: str) -> dict[str, object]:
    return {
        "id": resource_id,
        "name": name,
        "target": target,
        "description": f"{name} template",
        "path": f"C:/package/templates/{name}.{target}",
        "size_bytes": 123,
        "modified_ns": 456,
    }


def _template_payload(*, resources: list[dict[str, object]]) -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": True,
        "command": "resources list",
        "data": {"type": "templates", "resources": resources, "total": len(resources)},
        "error": None,
    }


def _template_show_payload(resource: dict[str, object]) -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": True,
        "command": "resources show",
        "data": {"type": "templates", "resource": resource},
        "error": None,
    }


def _required_resources() -> list[dict[str, object]]:
    from scripts.release import verify_packaged_cli

    return [
        _template_resource(
            resource_id=f"template.{Path(filename).suffix.removeprefix('.').casefold()}.{index:064x}",
            name=Path(filename).stem,
            target=Path(filename).suffix.removeprefix(".").casefold(),
        )
        for index, filename in enumerate(verify_packaged_cli._REQUIRED_TEMPLATE_FILES)
    ]


def test_packaged_cli_template_discovery_passes_returned_id_unchanged_to_convert(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    resources = _required_resources()
    template_id = str(
        min(
            (resource for resource in resources if resource["target"] == "docx"),
            key=lambda resource: str(resource["id"]),
        )["id"]
    )
    calls: list[tuple[str, ...]] = []

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        if args[:3] == ("resources", "list", "templates"):
            return subprocess.CompletedProcess(
                [str(binary_path), *args],
                0,
                stdout=json.dumps(_template_payload(resources=resources)),
                stderr="",
            )
        if args[:3] == ("resources", "show", "templates"):
            selected = next(resource for resource in resources if resource["id"] == args[3])
            return subprocess.CompletedProcess(
                [str(binary_path), *args],
                0,
                stdout=json.dumps(_template_show_payload(selected)),
                stderr="",
            )
        output = Path(args[args.index("--output") + 1])
        with zipfile.ZipFile(output, "w") as package:
            package.writestr(
                "word/document.xml",
                f"<document>{verify_packaged_cli._TEMPLATE_SMOKE_TEXT}</document>",
            )
        payload = {
            "protocol_version": 3,
            "success": True,
            "command": "convert",
            "data": {"output": str(output)},
            "error": None,
        }
        return subprocess.CompletedProcess(
            [str(binary_path), *args],
            0,
            stdout=json.dumps(payload),
            stderr="",
        )

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    output = verify_packaged_cli._run_template_resource_smoke(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)

    assert output == tmp_path / "canonical-template-id-smoke.docx"
    assert calls[0] == ("resources", "list", "templates", "--json", "--quiet")
    assert calls[1] == ("resources", "show", "templates", template_id, "--json", "--quiet")
    assert calls[2][calls[2].index("--template") + 1] == template_id


@pytest.mark.parametrize("mutation", ["extra_field", "value_mismatch"])
def test_packaged_cli_template_discovery_rejects_show_drift_from_list(
    mutation: str,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    resources = _required_resources()
    selected = min(
        (resource for resource in resources if resource["target"] == "docx"),
        key=lambda resource: str(resource["id"]),
    )
    shown = dict(selected)
    if mutation == "extra_field":
        shown["unexpected"] = True
    else:
        shown["description"] = "drifted description"

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        payload = (
            _template_payload(resources=resources)
            if args[:3] == ("resources", "list", "templates")
            else _template_show_payload(shown)
        )
        return subprocess.CompletedProcess(
            [str(binary_path), *args],
            0,
            stdout=json.dumps(payload),
            stderr="",
        )

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    expected = "resource contract is incomplete" if mutation == "extra_field" else "did not exactly match"
    with pytest.raises(RuntimeError, match=expected):
        verify_packaged_cli._run_template_resource_smoke(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


@pytest.mark.parametrize("mutation", ["missing_field", "duplicate_id", "target_mismatch"])
def test_packaged_cli_template_discovery_rejects_incomplete_or_ambiguous_resources(
    mutation: str,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    resources = _required_resources()
    if mutation == "missing_field":
        resources[0].pop("description")
    elif mutation == "duplicate_id":
        resources[1]["id"] = resources[0]["id"]
    else:
        resources[0]["id"] = f"template.xlsx.{'f' * 64}"
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda binary_path, *args, cwd: subprocess.CompletedProcess(
            [str(binary_path), *args],
            0,
            stdout=json.dumps(_template_payload(resources=resources)),
            stderr="",
        ),
    )

    expected = {
        "missing_field": "resource contract is incomplete",
        "duplicate_id": "duplicate canonical IDs",
        "target_mismatch": "resource value is invalid",
    }[mutation]
    with pytest.raises(RuntimeError, match=expected):
        verify_packaged_cli._run_template_resource_smoke(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)
