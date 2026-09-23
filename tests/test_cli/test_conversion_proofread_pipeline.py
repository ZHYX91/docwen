"""Real CLI/Application/Runtime coverage of generated-DOCX proofreading."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from docx import Document

from docwen_bundle.config_port import ConfigPortAdapter
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_cli.main import main
from docwen_runtime.config import ConfigLoader

pytestmark = [
    pytest.mark.integration,
    pytest.mark.pr_gate,
    pytest.mark.release_gate,
    pytest.mark.skipif(sys.platform == "darwin", reason="Primary document operations support Windows and Linux"),
]


@pytest.fixture
def invoke(tmp_path, monkeypatch, capsys):
    root = Path(__file__).resolve().parents[2]
    monkeypatch.setenv("DOCWEN_DATA_DIR", str(tmp_path / "profile"))
    loader = ConfigLoader(
        base_dir=root / "configs",
        user_dir=tmp_path / "profile" / "configs",
        runtime_overrides={"logger": {"console_enable": False}},
    )

    def run(*args):
        code = main(
            [*map(str, args), "--json", "--quiet"],
            runtime_port_factory=lambda: create_runtime_port(config_loader=loader),
            config_port_factory=lambda: ConfigPortAdapter(loader),
        )
        captured = capsys.readouterr()
        payload = json.loads(captured.out)
        assert code == 0, payload
        assert payload["success"], payload
        return payload

    return run, loader


def _docx(directory: Path) -> Path:
    outputs = list(directory.rglob("*.docx"))
    assert len(outputs) == 1
    return outputs[0]


def _content(path: Path):
    document = Document(path)
    return (
        [paragraph.text for paragraph in document.paragraphs],
        [comment.text for comment in document.comments],
        len(document.inline_shapes),
    )


@pytest.mark.parametrize("check", ["punct", "typo", "symbol", "sensitive", "all"])
def test_composed_output_matches_manual_convert_then_validate(tmp_path, invoke, check):
    run, _loader = invoke
    source = tmp_path / "校对.md"
    source.write_text(
        "# 校对\n\n公文才料（测试,秘密\n\n**加粗正文**\n\n"
        "| 列 | 内容 |\n| --- | --- |\n| 表格 | 保留 |\n\n"
        "```mermaid\nflowchart LR\nsubgraph 申请人\nA[提交]\nend\nA --> B[审批]\n```\n",
        encoding="utf-8",
    )
    original = source.read_bytes()
    plain = tmp_path / "plain"
    composed = tmp_path / "composed"
    manual = tmp_path / "manual.docx"

    run("convert", source, "--to", "docx", "--output-dir", plain)
    run("validate", _docx(plain), "--check", check, "--report", manual)
    run("convert", source, "--to", "docx", "--output-dir", composed, "--proofread", "--check", check)

    assert _content(_docx(composed)) == _content(manual)
    assert len(list(composed.iterdir())) == 1
    assert "_fromMd" in _docx(composed).parent.name
    assert source.read_bytes() == original


def test_disabled_config_still_publishes_generated_document(tmp_path, invoke):
    run, loader = invoke
    for name in ("symbol_pairing", "symbol_correction", "typos_rule", "sensitive_word"):
        loader.set_value(f"proofread.engine.enable_{name}", False)
    source = tmp_path / "disabled.md"
    source.write_text("# 保留正文\n\n没有启用校对也必须生成文档。\n", encoding="utf-8")
    destination = tmp_path / "output"

    run("convert", source, "--to", "docx", "--output-dir", destination, "--proofread")

    text, comments, _images = _content(_docx(destination))
    assert "没有启用校对也必须生成文档。" in text
    assert not comments


def test_batch_proofread_publishes_one_document_per_input(tmp_path, invoke):
    run, _loader = invoke
    sources = tmp_path / "inputs"
    sources.mkdir()
    for name in ("first", "second"):
        (sources / f"{name}.md").write_text(f"# {name}\n\n公文才料（\n", encoding="utf-8")
    output = tmp_path / "output"

    run(
        "batch",
        "convert",
        *sorted(sources.glob("*.md")),
        "--to",
        "docx",
        "--output-dir",
        output,
        "--proofread",
        "--check",
        "all",
    )

    assert len(list(output.rglob("*.docx"))) == 2
    assert len(list(output.iterdir())) == 2
