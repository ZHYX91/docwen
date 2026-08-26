from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_REPO_ROOT = Path(__file__).resolve().parents[2]


def _write_commented_docx(path: Path) -> None:
    document_xml = (
        f'<w:document xmlns:w="{_W_NS}"><w:body><w:p>'
        '<w:r><w:t>前</w:t></w:r><w:commentRangeStart w:id="0"/>'
        '<w:r><w:t>秘密</w:t></w:r><w:commentRangeEnd w:id="0"/>'
        "</w:p></w:body></w:document>"
    )
    comments_xml = (
        f'<w:comments xmlns:w="{_W_NS}"><w:comment w:id="0" w:author="DocWen">'
        "<w:p><w:r><w:t>敏感批注</w:t></w:r></w:p></w:comment></w:comments>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document_xml.encode())
        archive.writestr("word/comments.xml", comments_xml.encode())


def test_all_maintained_anchor_tools_bind_the_canonical_proofread_module() -> None:
    current_only_paths = (
        _REPO_ROOT / "tools" / "docx_spell_anchor_report.py",
        _REPO_ROOT / "tools" / "docx_spell_anchor_matrix_check.py",
    )

    for path in current_only_paths:
        source = path.read_text(encoding="utf-8")
        assert "docwen_plugin_proofread.anchor_report" in source
        assert "docwen.docx_spell" not in source

    parity_probe = (_REPO_ROOT / "tools" / "validation" / "probe_proofread_numbering_parity.py").read_text(
        encoding="utf-8"
    )
    augment_anchor_evidence = parity_probe[parity_probe.index("def _augment_anchor_evidence") :]
    assert "from docwen_plugin_proofread.anchor_report import" in augment_anchor_evidence


@pytest.mark.parametrize(
    "tool",
    ["docx_spell_anchor_report.py", "docx_spell_anchor_matrix_check.py"],
)
def test_anchor_tools_start_and_render_help(tool: str) -> None:
    completed = subprocess.run(
        [sys.executable, str(_REPO_ROOT / "tools" / tool), "--help"],
        cwd=_REPO_ROOT,
        check=False,
        capture_output=True,
    )

    assert completed.returncode == 0, completed.stderr.decode(errors="replace")
    assert b"--context-chars" in completed.stdout
    assert b"--redact" in completed.stdout


def test_anchor_report_cli_generates_a_redacted_rich_report(tmp_path: Path) -> None:
    source = tmp_path / "commented.docx"
    output = tmp_path / "anchor-report.md"
    _write_commented_docx(source)

    completed = subprocess.run(
        [
            sys.executable,
            str(_REPO_ROOT / "tools" / "docx_spell_anchor_report.py"),
            str(source),
            "--out",
            str(output),
            "--context-chars",
            "5",
            "--redact",
        ],
        cwd=_REPO_ROOT,
        check=False,
        capture_output=True,
    )

    assert completed.returncode == 0, completed.stderr.decode(errors="replace")
    report = output.read_text(encoding="utf-8")
    assert "# DOCX 批注锚点报告" in report
    assert "范围：`[1,3)`" in report
    assert "秘密" not in report
    assert "敏感批注" not in report
    assert "[██]" in report


def test_anchor_matrix_cli_passes_all_cases_with_output_local_runtime_logs(tmp_path: Path) -> None:
    output = tmp_path / "matrix"
    forbidden_log_dir = tmp_path / "outer-log-dir"
    environment = os.environ.copy()
    environment["DOCWEN_LOG_DIR"] = str(forbidden_log_dir)

    completed = subprocess.run(
        [
            sys.executable,
            str(_REPO_ROOT / "tools" / "docx_spell_anchor_matrix_check.py"),
            "--out-dir",
            str(output),
            "--keep",
        ],
        cwd=_REPO_ROOT,
        env=environment,
        check=False,
        capture_output=True,
    )

    assert completed.returncode == 0, completed.stderr.decode(errors="replace")
    index = (output / "index.md").read_text(encoding="utf-8")
    assert index.count("**PASS**") == 4
    assert "**FAIL**" not in index
    assert not forbidden_log_dir.exists()
    for case_dir in sorted(output.glob("case-*")):
        assert (case_dir / "runtime-logs" / "logs" / "docwen.log").is_file()
