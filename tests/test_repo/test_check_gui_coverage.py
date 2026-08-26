from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _load_module():
    project_root = Path(__file__).resolve().parents[2]
    script_path = project_root / "tools" / "check_gui_coverage.py"
    spec = importlib.util.spec_from_file_location("docwen_check_gui_coverage", script_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _write_coverage_xml(tmp_path: Path, body: str) -> Path:
    xml_path = tmp_path / "coverage-gui.xml"
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


def test_check_gui_coverage_reports_combined_gui_summary(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        "\n".join(
            [
                _class_xml("packages/apps/gui/src/docwen_gui/widgets/info_area.py", [1, 1, 0, 1]),
                _class_xml("packages/apps/gui/src/docwen_gui/app.py", [1, 0]),
                _class_xml("src/docwen/converter/md2docx/core.py", [1, 1, 1]),
            ]
        ),
    )

    exit_code = module.main([str(xml_path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "gui/: 66.67% (4/6 lines, 2 files)" in captured.out


def test_check_gui_coverage_fails_when_gui_files_are_missing(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        _class_xml("src/docwen/converter/md2docx/core.py", [1, 1, 0]),
    )

    exit_code = module.main([str(xml_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "[missing] gui/: no matching files found in coverage XML" in captured.out


def test_check_gui_coverage_supports_optional_threshold(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    module = _load_module()
    xml_path = _write_coverage_xml(
        tmp_path,
        _class_xml("packages/apps/gui/src/docwen_gui/widgets/info_area.py", [1, 0]),
    )

    exit_code = module.main([str(xml_path), "--fail-under", "75"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "[threshold-failed] gui/: 50.00% < required 75.00%" in captured.err
