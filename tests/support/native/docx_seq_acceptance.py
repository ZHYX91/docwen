"""Create real converter outputs for manual Word/WPS SEQ and REF acceptance.

This helper uses the production converter with test context adapters. It does
not open an Office host, update fields, paginate, or certify native behavior.
"""

from __future__ import annotations

import argparse
import hashlib
import inspect
import json
import shutil
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config.document_styles import build_document_style_catalog

_SOURCE = """Table: Generated first ^generated-first

| A |
|---|
| 1 |

Table: Generated second ^generated-second

| B |
|---|
| 2 |

First reference: @[[#^generated-first|first]].
Second reference: @[[#^generated-second|second]].
"""


def _sequence(paragraph: Any, cached: str) -> None:
    for kind in ("begin", "separate"):
        field = OxmlElement("w:fldChar")
        field.set(qn("w:fldCharType"), kind)
        paragraph.add_run()._r.append(field)
        if kind == "begin":
            instruction = OxmlElement("w:instrText")
            instruction.set(qn("xml:space"), "preserve")
            instruction.text = " SEQ Table \\* ARABIC "
            paragraph.add_run()._r.append(instruction)
    paragraph.add_run(cached)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(end)


def field_snapshot(path: Path) -> dict[str, Any]:
    """Inspect saved instructions/caches without asking a host to calculate them."""
    document = Document(path)
    fields: list[dict[str, Any]] = []
    for ordinal, paragraph in enumerate(document.element.body.iter(qn("w:p")), 1):
        stack: list[dict[str, Any]] = []
        bookmarks = [node.get(qn("w:name")) for node in paragraph.iter(qn("w:bookmarkStart"))]
        for node in paragraph.iter():
            if node.tag == qn("w:fldChar"):
                kind = node.get(qn("w:fldCharType"))
                if kind == "begin":
                    stack.append({"instruction": "", "cached_result": "", "result_started": False})
                elif kind == "separate" and stack:
                    stack[-1]["result_started"] = True
                elif kind == "end" and stack:
                    field = stack.pop()
                    field.pop("result_started")
                    fields.append({"paragraph": ordinal, "bookmarks": bookmarks, **field})
            elif stack and node.tag == qn("w:instrText"):
                stack[-1]["instruction"] += node.text or ""
            elif stack and node.tag == qn("w:t") and stack[-1]["result_started"]:
                stack[-1]["cached_result"] += node.text or ""
            elif node.tag == qn("w:fldSimple"):
                fields.append(
                    {
                        "paragraph": ordinal,
                        "bookmarks": bookmarks,
                        "instruction": node.get(qn("w:instr"), ""),
                        "cached_result": "".join(item.text or "" for item in node.iter(qn("w:t"))),
                    }
                )
        if stack:
            raise ValueError("Fixture contains an unterminated paragraph field")
    return {"package_sha256": hashlib.sha256(path.read_bytes()).hexdigest(), "fields": fields}


def create_fixture(directory: Path) -> dict[str, Path]:
    """Write paired controls into one new, exclusively owned directory."""
    repository = Path(__file__).resolve().parents[3]
    converter_path = Path(inspect.getfile(MdToDocxConverter)).resolve()
    if not converter_path.is_relative_to(repository):
        raise RuntimeError("The converter must be imported from the fixture's own checkout")
    directory.mkdir(parents=True, exist_ok=False)
    source = directory / "source.md"
    source.write_text(_SOURCE, encoding="utf-8")
    catalog = build_document_style_catalog(
        {}, request_options={"locale": "en_US"}, locales_dir=repository / "i18n/locales"
    )
    outputs: dict[str, Path] = {}
    for name, native_fields in (("control", False), ("shared-counter", True)):
        template = Document()
        style_name = catalog.name_for("table_caption")
        template.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        for position, cached in (("before", "1"), ("after", "2")):
            if position == "after":
                template.add_paragraph("{{body}}")
            caption = template.add_paragraph(f"Template {position}: ", style=style_name)
            if native_fields:
                _sequence(caption, cached)
            else:
                caption.add_run(cached)
            template.add_table(rows=1, cols=1).cell(0, 0).text = f"Fixed {position}"
        template_path = directory / f"{name}-template.docx"
        template.save(str(template_path))
        staging = directory / f"{name}-staging"
        staging.mkdir()
        request = ConversionRequest(
            request_id=f"seq-acceptance-{name}",
            input_refs=[FileRef(path=str(source), format="markdown", category="document")],
            target_format="docx",
            options={
                "template_name": str(template_path),
                "locale": "en_US",
                "heading_merge_mode": "never",
                "markdown_extensions": {"input": {"captions_references": True}},
            },
            output_policy=OutputPolicy(),
        )
        context = FakeExecutionContext(
            request,
            FakeWorkspaceHandle(str(source), str(staging)),
            FakeConfigView(),
            FakeProgressSink(),
            FakeCancellationTokenView(),
            FakePluginLogger(),
            document_style_catalog=catalog,
        )
        result = MdToDocxConverter().convert(context)
        if not result.success:
            raise RuntimeError(f"Fixture conversion failed: {result.error}")
        output = directory / f"{name}-before-update.docx"
        shutil.copyfile(result.artifacts[0].staging_path, output)
        outputs[name] = output
    manifest = {
        "schema": "docwen.native.seq-acceptance.v1",
        "native_acceptance": "pending",
        "converter_module": str(converter_path),
        "expected_generated_caption_caches": ["1", "2"],
        "observed_before_update": {name: field_snapshot(path) for name, path in outputs.items()},
        "limitation": "Source conversion and XML observations only; no Office field update, pagination or printing performed.",
    }
    (directory / "before-update.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return outputs


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("create").add_argument("directory", type=Path)
    commands.add_parser("inspect").add_argument("document", type=Path)
    args = parser.parse_args()
    if args.command == "create":
        print(json.dumps({name: str(path) for name, path in create_fixture(args.directory).items()}, indent=2))
    else:
        print(json.dumps(field_snapshot(args.document), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
