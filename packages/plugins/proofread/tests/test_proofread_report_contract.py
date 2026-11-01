"""Machine-readable Markdown proofread report 2.0 contracts."""

from __future__ import annotations

import json
from hashlib import sha256
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.contract

_PROJECT_ROOT = Path(__file__).resolve().parents[4]


def _context(
    source: Path,
    staging: Path,
    *,
    options: dict[str, bool] | None = None,
    proofread_rules: Any = None,
    config_values: dict[str, Any] | None = None,
) -> Any:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    request = ConversionRequest(
        request_id="proofread-report-contract",
        input_refs=[FileRef(path=str(source), format="markdown", category="text")],
        target_format="markdown",
        action_name="validate",
        options=options or {},
        output_policy=OutputPolicy(),
    )
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(str(source), str(staging)),
        FakeConfigView(config_values),
        FakeProgressSink(),
        CancellationToken().view(),
        FakePluginLogger(),
        proofread_rules=proofread_rules,
    )


def _read_report(result: Any) -> dict[str, Any]:
    return json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))


def test_report_2_0_uses_original_bytes_and_unicode_code_point_ranges(tmp_path: Path) -> None:
    from docwen_core.models.proofread import ProofreadRules
    from docwen_plugin_proofread.md_validator import MarkdownValidator

    # BOM, non-BMP emoji, a combining mark, a ZWJ sequence, CRLF, LF, and
    # an unmatched symbol at EOF exercise every coordinate boundary.
    text = "\ufeff😀e\u0301👩\u200d💻己\r\n第二行１２\n结尾（"
    raw = text.encode("utf-8")
    source = tmp_path / "coordinates.md"
    source.write_bytes(raw)
    result = MarkdownValidator().convert(
        _context(
            source,
            tmp_path,
            options={
                "enable_symbol_pairing": True,
                "enable_symbol_correction": True,
                "enable_typos_rule": True,
                "enable_sensitive_word": False,
            },
            proofread_rules=ProofreadRules(
                symbol_pairs=(("（", "）"),),
                symbol_map={"1": ("１",), "2": ("２",)},
                typos_map={"已": ("己",)},
            ),
        )
    )

    assert result.success is True
    report = _read_report(result)
    assert report["schema"] == "docwen.proofread_report.v2"
    assert report["source"] == {
        "content_sha256": sha256(raw).hexdigest(),
        "encoding": "utf-8",
        "decode_errors": "replace",
    }
    assert report["location_contract"] == {
        "id": "docwen.proofread-text-range",
        "version": 1,
        "coordinate_system": "unicode_code_point",
        "offset_base": 0,
        "line_base": 0,
        "column_base": 0,
        "range_end": "exclusive",
    }

    typo = next(issue for issue in report["issues"] if issue["rule_key"] == "typo")
    assert typo["range"] == {
        "start": {"offset": 7, "line": 0, "column": 7},
        "end": {"offset": 8, "line": 0, "column": 8},
    }
    assert typo["matched_text"] == "己"
    assert not {"line", "col_start", "col_end"}.intersection(typo)
    assert typo["fix"] == {
        "kind": "replace_text",
        "replacement": "已",
        "applicable": True,
    }

    symbol = next(
        issue for issue in report["issues"] if issue["rule_key"] == "symbol_correct" and issue["matched_text"] == "１"
    )
    assert symbol["range"] == {
        "start": {"offset": 13, "line": 1, "column": 3},
        "end": {"offset": 14, "line": 1, "column": 4},
    }
    assert symbol["fix"]["replacement"] == "1"

    eof_issue = next(issue for issue in report["issues"] if issue["rule_key"] == "symbol_pair")
    assert eof_issue["matched_text"] == "（"
    assert eof_issue["range"] == {
        "start": {"offset": 18, "line": 2, "column": 2},
        "end": {"offset": 19, "line": 2, "column": 3},
    }
    assert "fix" not in eof_issue


def test_bom_is_masked_without_shifting_frontmatter_or_visible_text(tmp_path: Path) -> None:
    from docwen_core.models.proofread import ProofreadRules
    from docwen_plugin_proofread.md_validator import MarkdownValidator, _sanitize_markdown

    text = "\ufeff---\r\ntitle: （\r\n---\r\n我己。"
    sanitized = _sanitize_markdown(text)
    assert sanitized.original_text == text
    assert len(sanitized.sanitized_text) == len(text)
    assert sanitized.sanitized_text[0] == " "
    assert "title" not in sanitized.sanitized_text

    source = tmp_path / "bom-frontmatter.md"
    source.write_bytes(text.encode("utf-8"))
    result = MarkdownValidator().convert(
        _context(
            source,
            tmp_path,
            options={
                "enable_symbol_pairing": True,
                "enable_symbol_correction": False,
                "enable_typos_rule": True,
                "enable_sensitive_word": False,
            },
            proofread_rules=ProofreadRules(
                symbol_pairs=(("（", "）"),),
                typos_map={"已": ("己",)},
            ),
        )
    )
    report = _read_report(result)
    assert [issue["rule_key"] for issue in report["issues"]] == ["typo"]
    issue = report["issues"][0]
    start = text.index("己")
    assert issue["range"]["start"] == {"offset": start, "line": 3, "column": 1}
    assert issue["matched_text"] == text[start : start + 1]


def test_disabled_checks_emit_empty_report_for_replace_decoded_source(tmp_path: Path) -> None:
    from docwen_plugin_proofread.md_validator import MarkdownValidator

    raw = b"clean\xffsource"
    source = tmp_path / "disabled.md"
    source.write_bytes(raw)
    result = MarkdownValidator().convert(
        _context(
            source,
            tmp_path,
            options={
                "enable_symbol_pairing": False,
                "enable_symbol_correction": False,
                "enable_typos_rule": False,
                "enable_sensitive_word": False,
            },
        )
    )

    assert result.success is True
    assert len(result.artifacts) == 1
    report = _read_report(result)
    assert report["issues"] == []
    assert report["summary"] == {}
    assert report["source"]["content_sha256"] == sha256(raw).hexdigest()
    assert report["source"]["decode_errors"] == "replace"
    assert result.metrics.extra["proofread_report"] == report


def test_suggestion_text_never_creates_an_undeclared_fix() -> None:
    from docwen_plugin_proofread.md_validator import _build_report_issue, _sanitize_markdown
    from docwen_plugin_proofread.text_validator import TextError

    sanitized = _sanitize_markdown("己（敏感词")
    typo_without_replacement = TextError(0, 1, "己", "已", "Typo", "typo")
    pairing = TextError(1, 2, "（", "）", "Unmatched Symbol", "pairing")
    sensitive = TextError(2, 5, "敏感词", "replace me", "Sensitive Word", "sensitive")

    for error, key in (
        (typo_without_replacement, "typo"),
        (pairing, "symbol_pair"),
        (sensitive, "sensitive"),
    ):
        issue = _build_report_issue(sanitized, error, key)
        assert "fix" not in issue


def test_issue_range_and_error_text_must_describe_the_same_source_slice() -> None:
    from docwen_plugin_proofread.md_validator import _build_report_issue, _sanitize_markdown
    from docwen_plugin_proofread.text_validator import TextError

    sanitized = _sanitize_markdown("己")
    mismatched = TextError(0, 1, "已", "己", "Typo", "typo", replacement="己")

    with pytest.raises(ValueError, match="does not match its authoritative source range"):
        _build_report_issue(sanitized, mismatched, "typo")


def test_schema_file_freezes_report_and_fix_contract() -> None:
    schema_path = _PROJECT_ROOT / "contracts" / "schemas" / "docwen.proofread_report.v2.schema.json"
    schema = json.loads(schema_path.read_text(encoding="utf-8"))

    assert schema["properties"]["schema"] == {"const": "docwen.proofread_report.v2"}
    source = schema["properties"]["source"]
    assert source["required"] == ["content_sha256", "encoding", "decode_errors"]
    location = schema["properties"]["location_contract"]["properties"]
    assert location["coordinate_system"] == {"const": "unicode_code_point"}
    assert location["offset_base"] == {"const": 0}
    assert location["line_base"] == {"const": 0}
    assert location["column_base"] == {"const": 0}
    assert location["range_end"] == {"const": "exclusive"}
    assert schema["$defs"]["fix"]["properties"] == {
        "kind": {"const": "replace_text"},
        "replacement": {"type": "string"},
        "applicable": {"const": True},
    }


@pytest.mark.parametrize(
    "literal",
    [
        "`１２３`",
        "``１２３``",
        "```１２３```",
        "`` first ` １２３\nsecond ``",
        "````text\n```\n１２３\n````",
        "````text\n````invalid closer\n１２３\n````",
        "~~~text\n１２３\n~~~",
        "> ````text\n> ```\n> １２３\n> ````",
        "- ````text\n  ```\n  １２３\n  ````",
        "[visible ``１２３``](https://example.test/１２３)",
        "- ```text\n  １２３\n",
        "> This ``１２３\nstill code``",
        "- This ``１２３\nstill code``",
    ],
)
def test_code_literals_never_produce_fixes_and_prose_ranges_remain_authoritative(literal: str, tmp_path: Path) -> None:
    from docwen_core.models.proofread import ProofreadRules
    from docwen_plugin_proofread.md_validator import MarkdownValidator

    prefix = "\ufeff# Probe\r\n\r\n" + literal + "\r\n\r\n😀 prose "
    text = prefix + "１２３\r\n"
    raw = text.encode("utf-8")
    source = tmp_path / "protected.md"
    source.write_bytes(raw)
    result = MarkdownValidator().convert(
        _context(
            source,
            tmp_path,
            options={
                "enable_symbol_pairing": False,
                "enable_symbol_correction": True,
                "enable_typos_rule": False,
                "enable_sensitive_word": False,
            },
            proofread_rules=ProofreadRules(symbol_map={"1": ("１",), "2": ("２",), "3": ("３",)}),
        )
    )
    assert result.success
    report = _read_report(result)
    assert report["source"]["content_sha256"] == sha256(raw).hexdigest()
    assert len(report["issues"]) == 3
    for index, issue in enumerate(sorted(report["issues"], key=lambda item: item["range"]["start"]["offset"])):
        start = len(prefix) + index
        assert issue["matched_text"] == "１２３"[index] == text[start : start + 1]
        assert issue["range"]["start"] == {"offset": start, "line": prefix.count("\n"), "column": 8 + index}
        assert issue["range"]["end"] == {"offset": start + 1, "line": prefix.count("\n"), "column": 9 + index}
        assert issue["fix"] == {"kind": "replace_text", "replacement": str(index + 1), "applicable": True}
    assert source.read_bytes() == raw


@pytest.mark.parametrize(
    "body", [r"\`１２３\`", "[１２３](https://example.test/１２３)", "![１２３](picture-１２３.png)"]
)
def test_literal_escapes_and_link_labels_remain_proofreadable_without_syntax_pairing_errors(
    body: str, tmp_path: Path
) -> None:
    from docwen_core.models.proofread import ProofreadRules
    from docwen_plugin_proofread.md_validator import MarkdownValidator

    text = "# Probe\n\n" + body
    source = tmp_path / "visible.md"
    source.write_bytes(text.encode("utf-8"))
    result = MarkdownValidator().convert(
        _context(
            source,
            tmp_path,
            options={
                "enable_symbol_pairing": True,
                "enable_symbol_correction": True,
                "enable_typos_rule": False,
                "enable_sensitive_word": False,
            },
            proofread_rules=ProofreadRules(
                symbol_pairs=(("[", "]"),), symbol_map={"1": ("１",), "2": ("２",), "3": ("３",)}
            ),
        )
    )
    assert result.success
    issues = _read_report(result)["issues"]
    assert len(issues) == 3
    assert {issue["rule_key"] for issue in issues} == {"symbol_correct"}
    assert sorted(issue["range"]["start"]["offset"] for issue in issues) == [
        text.index("１２３") + index for index in range(3)
    ]


@pytest.mark.parametrize("structural_tables", [False, True])
def test_proofread_uses_the_frozen_table_dialect_for_code_ownership(structural_tables: bool, tmp_path: Path) -> None:
    from docwen_core.models.proofread import ProofreadRules
    from docwen_plugin_proofread.md_validator import MarkdownValidator

    text = "# Probe\n\n| A | B |\n| --- | --- |\n| ` | １２３ ` |\n\n１２３ prose\n"
    source = tmp_path / "dialect.md"
    source.write_bytes(text.encode("utf-8"))
    result = MarkdownValidator().convert(
        _context(
            source,
            tmp_path,
            options={
                "enable_symbol_pairing": False,
                "enable_symbol_correction": True,
                "enable_typos_rule": False,
                "enable_sensitive_word": False,
            },
            config_values={"conversion": {"markdown_extensions": {"input": {"structural_tables": structural_tables}}}},
            proofread_rules=ProofreadRules(symbol_map={"1": ("１",), "2": ("２",), "3": ("３",)}),
        )
    )
    assert result.success
    report = _read_report(result)
    positions = [text.rindex("１２３") + index for index in range(3)]
    if not structural_tables:
        positions = [text.index("１２３") + index for index in range(3)] + positions
    assert sorted(issue["range"]["start"]["offset"] for issue in report["issues"]) == positions
    assert report["source"]["content_sha256"] == sha256(source.read_bytes()).hexdigest()


@pytest.mark.parametrize("extended_headings", [None, False, True])
@pytest.mark.parametrize("level", [6, 7, 8, 9, 10])
def test_proofread_uses_the_frozen_heading_dialect_for_code_ownership(
    extended_headings: bool | None, level: int, tmp_path: Path
) -> None:
    from docwen_core.models.proofread import ProofreadRules
    from docwen_plugin_proofread.md_validator import MarkdownValidator

    text = f"`start\r\n{'#' * level} １２３ [link](https://example.test)\r\nend`\r\n\r\n１２３ prose\r\n"
    raw = text.encode("utf-8")
    source = tmp_path / "heading-dialect.md"
    source.write_bytes(raw)
    config_values = (
        None
        if extended_headings is None
        else {"conversion": {"markdown_extensions": {"input": {"extended_headings": extended_headings}}}}
    )
    result = MarkdownValidator().convert(
        _context(
            source,
            tmp_path,
            options={
                "enable_symbol_pairing": False,
                "enable_symbol_correction": True,
                "enable_typos_rule": False,
                "enable_sensitive_word": False,
            },
            config_values=config_values,
            proofread_rules=ProofreadRules(symbol_map={"1": ("１",), "2": ("２",), "3": ("３",)}),
        )
    )
    assert result.success
    report = _read_report(result)
    positions = [(text.rindex("１２３") + index, 4, index) for index in range(3)]
    if level == 6 or (extended_headings and level <= 9):
        positions = [(text.index("１２３") + index, 1, level + 1 + index) for index in range(3)] + positions
    issues = sorted(report["issues"], key=lambda item: item["range"]["start"]["offset"])
    assert len(issues) == len(positions)
    for issue, (offset, line, column) in zip(issues, positions, strict=True):
        assert issue["range"]["start"] == {"offset": offset, "line": line, "column": column}
        assert issue["range"]["end"] == {"offset": offset + 1, "line": line, "column": column + 1}
        assert issue["matched_text"] == text[offset]
        assert issue["fix"] == {
            "kind": "replace_text",
            "replacement": str("１２３".index(text[offset]) + 1),
            "applicable": True,
        }
    assert report["source"]["content_sha256"] == sha256(raw).hexdigest()
    assert source.read_bytes() == raw
