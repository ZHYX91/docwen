"""MarkdownValidator — Markdown text proofreading.

Reads a Markdown file, sanitizes code blocks / YAML frontmatter / links,
validates the remaining text, and outputs a JSON proofread report.
"""

from __future__ import annotations

import json
import re
from bisect import bisect_right
from hashlib import sha256
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ProofreadConverterContext
    from docwen_plugin_proofread.text_validator import TextError

# ── Regex patterns for Markdown sanitization ───────────────────────────

_FENCE_RE = re.compile(r"^[ \t]*(```+|~~~+)")
_IMAGE_LINK_RE = re.compile(r"!\[([^\]]*)\]\(([^)]*)\)")
_MD_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]*)\)")

_REPORT_SCHEMA = "docwen.proofread_report.v2"
_SOURCE_ENCODING = "utf-8"
_SOURCE_DECODE_ERRORS = "replace"
_LOCATION_CONTRACT = {
    "id": "docwen.proofread-text-range",
    "version": 1,
    "coordinate_system": "unicode_code_point",
    "offset_base": 0,
    "line_base": 0,
    "column_base": 0,
    "range_end": "exclusive",
}


class MarkdownValidator:
    """Validate a Markdown file and produce a JSON proofread report.

    The report includes authoritative Unicode-code-point ranges, error text,
    suggestions, and a summary grouped by rule key.
    """

    def convert(self, context: ProofreadConverterContext) -> ConversionResult:
        from docwen_core.models.artifact import (
            ARTIFACT_KIND_PRIMARY,
            ArtifactManifest,
        )
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionMetrics,
            ConversionResult,
        )
        from docwen_core.paths import input_stem
        from docwen_plugin_proofread._common import (
            file_size,
            new_artifact_id,
            request_source_format,
            resolve_proofread_options,
        )
        from docwen_plugin_proofread.rules import (
            DIAGNOSTIC_INVALID_INPUT,
            DIAGNOSTIC_OK,
            DIAGNOSTIC_SKIPPED,
        )
        from docwen_plugin_proofread.text_validator import TextValidator, rule_key

        task_id = context.request.request_id
        input_path = context.workspace.input_path
        input_stem_val = input_stem(input_path)
        source_format = request_source_format(context)
        source_display_name = (
            Path(context.request.input_refs[0].path).name if context.request.input_refs else Path(input_path).name
        )

        context.cancellation.check()

        # ── Resolve proofread options from config + request options ──────
        opts = resolve_proofread_options(context)
        symbol_pairing = opts["enable_symbol_pairing"]
        symbol_correction = opts["enable_symbol_correction"]
        typos_rule = opts["enable_typos_rule"]
        sensitive_word = opts["enable_sensitive_word"]

        proofread_rules = context.proofread_rules
        symbol_pairs = list(proofread_rules.symbol_pairs) if proofread_rules else None
        symbol_map = (
            {key: list(values) for key, values in proofread_rules.symbol_map.items()}
            if proofread_rules and proofread_rules.symbol_map
            else None
        )
        typos_map = (
            {key: list(values) for key, values in proofread_rules.typos_map.items()}
            if proofread_rules and proofread_rules.typos_map
            else None
        )
        sensitive_map = (
            {key: list(values) for key, values in proofread_rules.sensitive_words.items()}
            if proofread_rules and proofread_rules.sensitive_words
            else None
        )

        try:
            source_bytes = Path(input_path).read_bytes()
        except OSError as exc:
            msg = f"Cannot read input file: {exc}"
            context.logger.error(msg)
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="invalid_input",
                    message=msg,
                    diagnostic_code=DIAGNOSTIC_INVALID_INPUT,
                ),
                diagnostics=[ConversionDiagnostic(level="error", message=msg, code=DIAGNOSTIC_INVALID_INPUT)],
            )

        input_bytes = len(source_bytes)
        text = source_bytes.decode(_SOURCE_ENCODING, errors=_SOURCE_DECODE_ERRORS)

        # ── Sanitize Markdown ──────────────────────────────────────────
        sanitized = _sanitize_markdown(text)
        context.logger.info(f"Sanitized MD: original {len(text)} chars → {len(sanitized.sanitized_text)} chars")

        # ── Validate ───────────────────────────────────────────────────
        checks_enabled = {
            "symbol_pairing": symbol_pairing,
            "symbol_correction": symbol_correction,
            "typos_rule": typos_rule,
            "sensitive_word": sensitive_word,
        }
        checks_active = any(checks_enabled.values())
        if checks_active:
            validator = TextValidator(
                symbol_pairs=symbol_pairs,
                symbol_map=symbol_map,
                typos_map=typos_map,
                sensitive_words=sensitive_map,
                enabled=checks_enabled,
                lang="en",
            )
            errors = validator.validate_text(sanitized.sanitized_text)
        else:
            context.logger.info("No proofread checks enabled — producing an empty report")
            errors = []

        # ── Map offsets to line:col ────────────────────────────────────
        issues: list[dict] = []
        for err in errors:
            issues.append(_build_report_issue(sanitized, err, rule_key(err.source)))

        # ── Build summary ──────────────────────────────────────────────
        summary: dict[str, int] = {}
        for it in issues:
            summary[it["rule_key"]] = summary.get(it["rule_key"], 0) + 1

        report = {
            "schema": _REPORT_SCHEMA,
            "file": source_display_name,
            "source": {
                "content_sha256": sha256(source_bytes).hexdigest(),
                "encoding": _SOURCE_ENCODING,
                "decode_errors": _SOURCE_DECODE_ERRORS,
            },
            "location_contract": dict(_LOCATION_CONTRACT),
            "checks_enabled": checks_enabled,
            "issues": issues,
            "summary": summary,
        }

        context.progress.report_progress(80.0, "Writing proofread report")

        # ── Write JSON report to staging ───────────────────────────────
        report_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".json")
        Path(report_path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

        output_bytes = file_size(report_path)
        context.progress.report_progress(100.0, "Markdown proofread complete")

        artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=report_path,
            suggested_name=f"{input_stem_val}_proofread_report.json",
            media_type="application/json",
            metadata={
                "source_format": source_format,
                "issues_found": len(issues),
                "checks_enabled": report["checks_enabled"],
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)
        context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)

        diagnostic = ConversionDiagnostic(
            level="info",
            message=(
                f"Proofread complete: {len(issues)} issue(s) found. Summary: {summary}"
                if checks_active
                else "No proofread checks enabled; an empty report was produced."
            ),
            code=DIAGNOSTIC_OK if checks_active else DIAGNOSTIC_SKIPPED,
        )

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact],
            diagnostics=[diagnostic],
            metrics=ConversionMetrics(
                input_bytes=input_bytes,
                output_bytes=output_bytes,
                extra={
                    "source_format": source_format,
                    "issues_found": len(issues),
                    "issue_summary": summary,
                    "proofread_report": report,
                },
            ),
        )


# ── Markdown sanitization ──────────────────────────────────────────────


class _SanitizedMarkdown:
    """Holds sanitized text with original line-start offsets."""

    __slots__ = ("line_starts", "original_text", "sanitized_text")

    def __init__(self, original_text: str, sanitized_text: str) -> None:
        self.original_text = original_text
        self.sanitized_text = sanitized_text
        self.line_starts = _compute_line_starts(original_text)

    def offset_to_contract_line_col(self, offset: int) -> tuple[int, int]:
        """Return a zero-based Unicode-code-point line and column.

        ``offset`` may equal ``len(original_text)`` so an exclusive range end
        at EOF remains representable without clamping it to the last character.
        """
        offset = max(0, min(offset, len(self.original_text)))
        idx = bisect_right(self.line_starts, offset) - 1
        if idx < 0:
            return 0, offset
        line_start = self.line_starts[idx]
        return idx, offset - line_start


def _build_report_issue(sanitized: _SanitizedMarkdown, error: TextError, issue_rule_key: str) -> dict:
    """Project one validator issue into the schema-2.0 location contract."""
    start = int(error.start_pos)
    end = int(error.end_pos)
    text_length = len(sanitized.original_text)
    if start < 0 or end <= start or end > text_length:
        raise ValueError(f"Invalid proofread issue range [{start}, {end}) for text length {text_length}")

    start_line, start_column = sanitized.offset_to_contract_line_col(start)
    end_line, end_column = sanitized.offset_to_contract_line_col(end)
    matched_text = sanitized.original_text[start:end]
    error_text = str(error.error_text)
    if matched_text != error_text:
        raise ValueError(
            "Proofread issue text does not match its authoritative source range: "
            f"range=[{start}, {end}); matched={matched_text!r}; error_text={error_text!r}"
        )

    source = str(error.source)
    issue = {
        "range": {
            "start": {"offset": start, "line": start_line, "column": start_column},
            "end": {"offset": end, "line": end_line, "column": end_column},
        },
        "matched_text": matched_text,
        "error_text": error_text,
        "suggestion": str(error.suggestion),
        "error_type": str(error.error_type),
        "source": source,
        "rule_key": issue_rule_key,
    }
    replacement = error.replacement
    if source in {"typo", "symbol"} and isinstance(replacement, str):
        issue["fix"] = {
            "kind": "replace_text",
            "replacement": replacement,
            "applicable": True,
        }
    return issue


def _sanitize_markdown(text: str) -> _SanitizedMarkdown:
    """Remove code blocks, YAML frontmatter, and inline code/links from *text*.

    Keeps newlines so line numbers stay correct.
    """
    text = text or ""
    body_offset = 1 if text.startswith("﻿") else 0
    body = text[body_offset:]

    spans: list[tuple[int, int]] = []
    spans.extend((start + body_offset, end + body_offset) for start, end in _find_yaml_front_matter_span(body))
    spans.extend((start + body_offset, end + body_offset) for start, end in _find_fenced_code_spans(body))
    spans = _merge_spans(spans)

    chars = list(text)
    if body_offset:
        chars[0] = " "

    for start, end in spans:
        _blank_range_keep_newlines(chars, start, end)

    _blank_inline_code_and_links(chars, text, spans)
    _blank_markdown_escaped_pairing_symbols(chars, text)

    sanitized = "".join(chars)
    return _SanitizedMarkdown(original_text=text, sanitized_text=sanitized)


def _blank_markdown_escaped_pairing_symbols(chars: list[str], original: str) -> None:
    """Blank escaped symmetric quotes without changing source offsets.

    Escaped brackets remain visible punctuation in rendered Markdown and must
    keep participating in pairing.  A backslash before a quote, however, is
    an explicit request to treat that quote as a literal rather than as a
    delimiter for the proofreader.
    """

    pairing_symbols = frozenset("'\"")
    for index, char in enumerate(original):
        if char not in pairing_symbols or index == 0:
            continue
        backslash_count = 0
        cursor = index - 1
        while cursor >= 0 and original[cursor] == "\\":
            backslash_count += 1
            cursor -= 1
        if backslash_count % 2:
            chars[index] = " "


# ── Line position helpers ──────────────────────────────────────────────


def _compute_line_starts(text: str) -> list[int]:
    starts = [0]
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == "\r":
            if i + 1 < n and text[i + 1] == "\n":
                starts.append(i + 2)
                i += 2
                continue
            starts.append(i + 1)
            i += 1
            continue
        if ch == "\n":
            starts.append(i + 1)
        i += 1
    return starts


# ── YAML frontmatter ───────────────────────────────────────────────────


def _find_yaml_front_matter_span(text: str) -> list[tuple[int, int]]:
    if not text.startswith("---"):
        return []
    if not _line_is_front_matter_delimiter(text, 0):
        return []
    for line_start, line_end in _iter_line_spans(text):
        if line_start == 0:
            continue
        if _line_is_front_matter_delimiter(text, line_start):
            return [(0, line_end)]
    return [(0, len(text))]


def _line_is_front_matter_delimiter(text: str, line_start: int) -> bool:
    line = text[line_start:]
    if line.startswith("---"):
        i = line_start + 3
        while i < len(text) and text[i] not in "\r\n":
            if not text[i].isspace():
                return False
            i += 1
        return True
    return False


# ── Fenced code blocks ─────────────────────────────────────────────────


def _find_fenced_code_spans(text: str) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    current_fence: str | None = None
    current_start = 0

    for line_start, line_end in _iter_line_spans(text):
        line = text[line_start:line_end]
        m = _FENCE_RE.match(line)
        if not m:
            continue
        fence = m.group(1)
        if current_fence is None:
            current_fence = fence
            current_start = line_start
            continue
        if fence and current_fence and fence[0] == current_fence[0]:
            spans.append((current_start, line_end))
            current_fence = None

    if current_fence is not None:
        spans.append((current_start, len(text)))

    return spans


# ── Inline code and links ──────────────────────────────────────────────


def _blank_inline_code_and_links(chars: list[str], original: str, spans: list[tuple[int, int]]) -> None:
    for line_start, line_end in _iter_line_spans(original):
        if _pos_in_spans(line_start, spans):
            continue
        _blank_inline_code_in_range(chars, line_start, line_end)
        _blank_links_in_range(chars, original, line_start, line_end)


def _blank_inline_code_in_range(chars: list[str], start: int, end: int) -> None:
    open_tick: int | None = None
    for i in range(start, end):
        if chars[i] in ("\n", "\r"):
            break
        if chars[i] != "`":
            continue
        if open_tick is None:
            open_tick = i
        else:
            _blank_range_keep_newlines(chars, open_tick, i + 1)
            open_tick = None


def _blank_links_in_range(chars: list[str], original: str, start: int, end: int) -> None:
    segment = original[start:end]
    for m in list(_IMAGE_LINK_RE.finditer(segment)):
        _blank_link_match(chars, original, start, m)
    for m in list(_MD_LINK_RE.finditer(segment)):
        if m.start() > 0 and segment[m.start() - 1] == "!":
            continue
        _blank_link_match(chars, original, start, m)


def _blank_link_match(chars: list[str], original: str, base: int, match: re.Match[str]) -> None:
    m0s, m0e = match.span(0)
    g1s, g1e = match.span(1)
    abs_m0s = base + m0s
    abs_m0e = base + m0e
    abs_g1s = base + g1s
    abs_g1e = base + g1e
    _blank_range_keep_newlines(chars, abs_m0s, abs_m0e)
    for i in range(abs_g1s, abs_g1e):
        chars[i] = original[i]


# ── Generic helpers ────────────────────────────────────────────────────


def _iter_line_spans(text: str):
    pos = 0
    for line in text.splitlines(keepends=True):
        start = pos
        pos += len(line)
        yield start, pos
    if text and not text.endswith(("\n", "\r")):
        return


def _merge_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    if not spans:
        return []
    spans = sorted(spans)
    merged: list[tuple[int, int]] = []
    cur_s, cur_e = spans[0]
    for s, e in spans[1:]:
        if s <= cur_e:
            cur_e = max(cur_e, e)
        else:
            merged.append((cur_s, cur_e))
            cur_s, cur_e = s, e
    merged.append((cur_s, cur_e))
    return merged


def _blank_range_keep_newlines(chars: list[str], start: int, end: int) -> None:
    n = len(chars)
    start = max(0, min(start, n))
    end = max(0, min(end, n))
    for i in range(start, end):
        if chars[i] in ("\n", "\r"):
            continue
        chars[i] = " "


def _pos_in_spans(pos: int, spans: list[tuple[int, int]]) -> bool:
    for s, e in spans:
        if s <= pos < e:
            return True
        if pos < s:
            return False
    return False
