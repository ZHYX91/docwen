"""Proofread reports keep the actual paragraph, list and Setext boundaries."""

from __future__ import annotations

from hashlib import sha256
from pathlib import Path

import pytest

from docwen_core.models.proofread import ProofreadRules
from docwen_plugin_proofread.md_validator import MarkdownValidator

from .test_proofread_report_contract import _context, _read_report

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("eol", ["\n", "\r\n"])
@pytest.mark.parametrize(
    ("body", "visible_occurrences"),
    [
        pytest.param(
            "start `open\n-\n[１２３](https://example.test) close`\n",
            (0,),
            id="setext-closes-the-inline-scope",
        ),
        pytest.param(
            "start `open\n１２３ close`\n===\n\n１２３ prose\n",
            (1,),
            id="multiline-setext-keeps-its-own-code",
        ),
        pytest.param(
            "start `open\n2. １２３\nclose`\n\n１２３ prose\n",
            (1,),
            id="noninitial-ordered-marker-stays-in-paragraph",
        ),
        pytest.param(
            "start `open\n1. １２３\nclose`\n\n１２３ prose\n",
            (0, 1),
            id="initial-ordered-marker-interrupts-paragraph",
        ),
        pytest.param(
            "2. start `open\n3. １２３ close`\n\n１２３ prose\n",
            (0, 1),
            id="real-ordered-siblings-have-separate-inline-scopes",
        ),
        pytest.param(
            "- text\nlazy\n  ```\n  １２３\n\n# Real\n１２３ prose\n",
            (1,),
            id="lazy-paragraph-keeps-list-fence-owner",
        ),
        pytest.param(
            "- text\nlazy\n  <pre>\n  １２３\n\n# Real\n１２３ prose\n",
            (1,),
            id="lazy-paragraph-keeps-list-html-owner",
        ),
        pytest.param(
            "- outer\n  - text\nlazy\n    ```\n    １２３\n\n# Real\n１２３ prose\n",
            (1,),
            id="nested-lazy-paragraph-keeps-fence-owner",
        ),
        pytest.param(
            "- ```text\n  literal\n１２３ prose\n",
            (),
            id="list-fence-keeps-its-unseparated-dedented-body",
        ),
        pytest.param(
            "- <pre>\n  literal\n１２３ prose\n",
            (),
            id="list-html-keeps-its-unseparated-dedented-body",
        ),
        pytest.param(
            "- ```text\n  literal\n\n１２３ prose\n",
            (0,),
            id="fence-body-does-not-enable-lazy-paragraph-after-blank",
        ),
        pytest.param(
            "- <pre>\n  literal\n\n１２３ prose\n",
            (0,),
            id="html-body-does-not-enable-lazy-paragraph-after-blank",
        ),
        pytest.param(
            "> start `open\n-\n１２３ close`\n",
            (0,),
            id="top-level-empty-list-does-not-extend-quote-code",
        ),
        pytest.param(
            "start `open\n***\n１２３ close`\n",
            (0,),
            id="thematic-break-ends-inline-scope",
        ),
    ],
)
def test_report_fixes_only_visible_text_in_its_actual_block(
    body: str, visible_occurrences: tuple[int, ...], eol: str, tmp_path: Path
) -> None:
    # BOM and a non-BMP character precede each case without merging its first
    # paragraph. Expected offsets come only from the original authored text.
    text = ("\ufeff😀\n\n" + body).replace("\n", eol)
    raw = text.encode("utf-8")
    source = tmp_path / "blocks.md"
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
    assert result.success, result.error
    report = _read_report(result)
    starts = [index for index in range(len(text)) if text.startswith("１２３", index)]
    expected_offsets = [starts[occurrence] + digit for occurrence in visible_occurrences for digit in range(3)]
    issues = sorted(report["issues"], key=lambda issue: issue["range"]["start"]["offset"])
    assert [issue["range"]["start"]["offset"] for issue in issues] == expected_offsets
    for issue, offset in zip(issues, expected_offsets, strict=True):
        line = text[:offset].count("\n")
        column = offset - (text.rfind("\n", 0, offset) + 1)
        assert issue["rule_key"] == "symbol_correct"
        assert issue["matched_text"] == text[offset : offset + 1]
        assert issue["range"]["start"] == {"offset": offset, "line": line, "column": column}
        assert issue["range"]["end"] == {"offset": offset + 1, "line": line, "column": column + 1}
        assert issue["fix"] == {
            "kind": "replace_text",
            "replacement": str("１２３".index(text[offset]) + 1),
            "applicable": True,
        }
    assert report["source"]["content_sha256"] == sha256(raw).hexdigest()
    assert source.read_bytes() == raw
