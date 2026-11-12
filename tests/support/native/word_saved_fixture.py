"""Prepare synthetic current-format DOCX for native Word save regression fixtures.

The generated file is not a Word-saved fixture. A native Word edit/save and
independent readback must follow before replacing the checked-in fixture.
"""

from __future__ import annotations

import argparse
import hashlib
import inspect
import json
import shutil
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from lxml import etree
from PIL import Image
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config.document_styles import build_document_style_catalog

REPLACEMENT_SOURCE = "待 Word 替换的普通正文。"
REPLACEMENT_TEXT = "Word 实测追加：从独立 DOCX 的当前内容回转。"
SOURCE = """# 合成回归文档

待 Word 替换的普通正文。

**粗体正文**与*斜体正文*应保留。

Table: 指标 ^metrics

| 项目 | 数值 |
| --- | --- |
| 甲 | 12 |
| 乙 | ^ |

见 @[[#Table: 指标]]。

####### 深层标题

注释[^endnote:n]。

[^endnote:n]: 这是尾注正文。

Figure: 合成图 ^figure

![合成图](synthetic.png)
"""


def _context(directory: Path, source: Path, target: str, direction: str) -> FakeExecutionContext:
    staging = directory / f"staging-{target}"
    staging.mkdir()
    reference = FileRef(path=str(source), format="markdown" if target == "docx" else "docx", category="document")
    request = ConversionRequest(
        request_id=f"word-fixture-{target}",
        input_refs=[reference],
        target_format=target,
        options={
            "markdown_extensions": {direction: MarkdownExtensions.obsidian().to_dict()},
            "heading_merge_mode": "never",
            "remove_numbering": False,
        },
        output_policy=OutputPolicy(),
    )
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(str(source), str(staging), (reference,)),
        FakeConfigView(),
        FakeProgressSink(),
        FakeCancellationTokenView(),
        FakePluginLogger(),
        document_style_catalog=build_document_style_catalog(
            {"gui": {"language": {"locale": "zh_CN"}}},
            locales_dir=Path(__file__).resolve().parents[3] / "i18n/locales",
        ),
    )


def package_snapshot(path: Path) -> dict[str, Any]:
    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(path) as package:
        document = etree.fromstring(package.read("word/document.xml"))
        controls = []
        for control in document.xpath("//w:sdt", namespaces=namespaces):
            controls.append(
                {
                    "tag": control.xpath("string(w:sdtPr/w:tag/@w:val)", namespaces=namespaces),
                    "blocks": len(control.find("w:sdtContent", namespaces)),
                    "text": "".join(control.xpath("w:sdtContent//w:t/text()", namespaces=namespaces)),
                }
            )
        return {
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "bytes": path.stat().st_size,
            "content_controls": controls,
            "custom_xml": sorted(name for name in package.namelist() if name.startswith("customXml/")),
        }


def create_fixture(directory: Path) -> Path:
    repository = Path(__file__).resolve().parents[3]
    converter = Path(inspect.getfile(MdToDocxConverter)).resolve()
    if not converter.is_relative_to(repository):
        raise RuntimeError("Converter must come from this checkout")
    directory.mkdir(parents=True, exist_ok=False)
    source = directory / "source.md"
    source.write_text(SOURCE, encoding="utf-8")
    Image.new("RGB", (80, 40), (30, 100, 180)).save(directory / "synthetic.png")
    converted = MdToDocxConverter().convert(_context(directory, source, "docx", "input"))
    if not converted.success:
        raise RuntimeError(f"Fixture conversion failed: {converted.error}")
    output = directory / "current-format-before-word.docx"
    shutil.copyfile(converted.artifacts[0].staging_path, output)
    restored = DocxToMarkdownConverter().convert(_context(directory, output, "md", "output"))
    if not restored.success:
        raise RuntimeError(f"Source readback failed: {restored.error}")
    markdown = Path(restored.artifacts[0].staging_path).read_text(encoding="utf-8")
    for expected in (
        REPLACEMENT_SOURCE,
        "**粗体正文**与*斜体正文*应保留。",
        "| 甲 | 12 |",
        "| 乙 | ^ |",
        "@[[#Table: 指标]]",
        "####### 深层标题",
        "[^endnote:1]",
        "这是尾注正文。",
    ):
        if expected not in markdown:
            raise RuntimeError(f"Missing baseline expectation: {expected}")
    (directory / "source-readback.md").write_text(markdown, encoding="utf-8")
    manifest = {
        "schema": "docwen.native.word-save-preparation.v1",
        "native_acceptance": "pending",
        "converter_module": str(converter),
        "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "output": package_snapshot(output),
        "native_edit": {"find": REPLACEMENT_SOURCE, "replace": REPLACEMENT_TEXT},
        "negative_edit": "Add one ordinary paragraph inside the final generated Figure content control; preserve its caption and image.",
        "limitation": "Prepared and read with source converters only. No native Word edit or save has occurred.",
    }
    (directory / "preparation.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    print(create_fixture(parser.parse_args().directory))


if __name__ == "__main__":
    main()
