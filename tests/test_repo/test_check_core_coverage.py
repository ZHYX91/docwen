from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _load_module():
    project_root = Path(__file__).resolve().parents[2]
    script_path = project_root / "tools" / "check_core_coverage.py"
    spec = importlib.util.spec_from_file_location("docwen_check_core_coverage", script_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _write_coverage_xml(tmp_path: Path, body: str) -> Path:
    xml_path = tmp_path / "coverage.xml"
    xml_path.write_text(
        f"""<?xml version="1.0" ?>
<coverage>
  <packages>
    <package name="docwen">
      <classes>
{body}
      </classes>
    </package>
  </packages>
</coverage>
""",
        encoding="utf-8",
    )
    return xml_path


def _class_xml(filename: str, hits: list[int]) -> str:
    line_nodes = "\n".join(
        f'            <line number="{index}" hits="{hit}" />' for index, hit in enumerate(hits, start=1)
    )
    return f"""
        <class name="{Path(filename).name}" filename="{filename}">
          <lines>
{line_nodes}
          </lines>
        </class>"""


def test_check_core_coverage_reports_plan_domains(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        "\n".join(
            [
                _class_xml("packages/core/src/docwen_core/models/result.py", [1, 1, 0, 1]),
                _class_xml("packages/application/src/docwen_application/controller.py", [1, 1, 0, 1]),
                _class_xml("packages/runtime/src/docwen_runtime/engine/task_manager.py", [1, 1]),
            ]
        ),
    )

    exit_code = module.main([str(xml_path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "core/: 75.00% (3/4 lines, 1 files)" in captured.out
    assert "application/: 75.00% (3/4 lines, 1 files)" in captured.out
    assert "runtime/: 100.00% (2/2 lines, 1 files)" in captured.out
    assert "combined core domains: 80.00% (8/10 lines)" in captured.out


def test_check_core_coverage_fails_when_required_domain_is_missing(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        _class_xml("packages/core/src/docwen_core/models/result.py", [1, 1, 0]),
    )

    exit_code = module.main([str(xml_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "[missing] application/: no matching files found in coverage XML" in captured.out
    assert "[missing] runtime/: no matching files found in coverage XML" in captured.out


def test_check_core_coverage_supports_optional_thresholds(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        "\n".join(
            [
                _class_xml("packages/core/src/docwen_core/models/result.py", [1, 0]),
                _class_xml("packages/application/src/docwen_application/controller.py", [1]),
                _class_xml("packages/runtime/src/docwen_runtime/engine/task_manager.py", [1]),
            ]
        ),
    )

    exit_code = module.main([str(xml_path), "--fail-under", "core/=80"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "[threshold-failed] core/: 50.00% < required 80.00%" in captured.err


def test_check_core_coverage_soft_gate_reports_without_failing(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        "\n".join(
            [
                _class_xml("packages/core/src/docwen_core/models/result.py", [1, 0]),
                _class_xml("packages/application/src/docwen_application/controller.py", [1]),
                _class_xml("packages/runtime/src/docwen_runtime/engine/task_manager.py", [1]),
            ]
        ),
    )

    exit_code = module.main([str(xml_path), "--report-under", "core/=80", "--report-under", "combined=85"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "[report-threshold] core/: 50.00% < required 80.00%" in captured.err
    assert "[report-threshold] combined: 75.00% < required 85.00%" in captured.err


def test_check_core_coverage_fail_under_supports_combined_threshold(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        "\n".join(
            [
                _class_xml("packages/core/src/docwen_core/models/result.py", [1, 0]),
                _class_xml("packages/application/src/docwen_application/controller.py", [1]),
                _class_xml("packages/runtime/src/docwen_runtime/engine/task_manager.py", [1]),
            ]
        ),
    )

    exit_code = module.main([str(xml_path), "--fail-under", "combined=85"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "[threshold-failed] combined: 75.00% < required 85.00%" in captured.err


def test_check_core_coverage_soft_gate_uses_built_in_thresholds(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        "\n".join(
            [
                _class_xml("packages/core/src/docwen_core/models/result.py", [1]),
                _class_xml("packages/application/src/docwen_application/controller.py", [1]),
                _class_xml("packages/runtime/src/docwen_runtime/engine/task_manager.py", [1]),
            ]
        ),
    )

    exit_code = module.main([str(xml_path), "--soft-gate"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "[report-threshold] combined: 100.00% < required 55.00%" not in captured.err


def test_check_core_coverage_emits_github_actions_annotations(
    tmp_path: Path, capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        "\n".join(
            [
                _class_xml("packages/core/src/docwen_core/models/result.py", [1, 0]),
                _class_xml("packages/application/src/docwen_application/controller.py", [1]),
                _class_xml("packages/runtime/src/docwen_runtime/engine/task_manager.py", [1]),
            ]
        ),
    )
    monkeypatch.setenv("GITHUB_ACTIONS", "true")

    exit_code = module.main([str(xml_path), "--report-under", "core/=80", "--fail-under", "combined=90"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "::warning::core-domain-coverage core/ 50.00% < 80.00%" in captured.out
    assert "::error::core-domain-coverage combined 75.00% < 90.00%" in captured.out
