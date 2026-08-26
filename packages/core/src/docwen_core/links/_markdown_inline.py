"""Small CommonMark-aware scanners used by request-scoped link processing."""

from __future__ import annotations

import re
import string
from contextvars import ContextVar
from dataclasses import dataclass
from urllib.parse import quote

_ASCII_PUNCTUATION = frozenset(string.punctuation)
_ASCII_WHITESPACE = " \t\r\n"
_TITLE_RE = re.compile(
    r"""(?:"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*')""",
    re.DOTALL,
)
_IMAGE_SIZE_RE = re.compile(r"=\s*\d*x\d*")
_ACTIVE_PROTECTED_TOKENS: ContextVar[frozenset[str]] = ContextVar(
    "docwen_active_markdown_tokens",
    default=frozenset(),
)
_BLOCK_INTERRUPT_RE = re.compile(
    r"^[ \t]{0,3}(?:"
    r"#{1,9}[ \t]+|"
    r">|"
    r"(?:[-+*]|0{0,8}1[.)])[ \t]+|"
    r"`{3,}(?![^\r\n]*`)|~{3,}"
    r")"
)
_HTML_BLOCK_INTERRUPT_RE = re.compile(
    r"^[ \t]{0,3}(?:"
    r"<!--|<\?|<![A-Z]|<!\[CDATA\[|"
    r"</?(?:address|article|aside|base|basefont|blockquote|body|caption|"
    r"center|col|colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|"
    r"figcaption|figure|footer|form|frame|frameset|h[1-6]|head|header|"
    r"hr|html|iframe|legend|li|link|main|menu|menuitem|nav|noframes|ol|"
    r"meta|optgroup|option|p|param|pre|script|section|source|style|summary|"
    r"table|tbody|td|textarea|tfoot|th|thead|title|tr|track|ul)"
    r"(?:[ \t/>]|$)"
    r")",
    re.IGNORECASE,
)


@dataclass(frozen=True, slots=True)
class MarkdownInlineLink:
    """A parsed inline Markdown link or image at one source position."""

    end: int
    label: str
    target: str
    is_image: bool


@dataclass(frozen=True, slots=True)
class MarkdownDestination:
    """Validated destination plus its exact title/size suffix."""

    destination: str
    suffix: str


def _is_backslash_escaped(text: str, index: int) -> bool:
    backslashes = 0
    cursor = index - 1
    while cursor >= 0 and text[cursor] == "\\":
        backslashes += 1
        cursor -= 1
    return backslashes % 2 == 1


def _contains_active_protected_token(text: str) -> bool:
    return any(token in text for token in _ACTIVE_PROTECTED_TOKENS.get())


def _has_only_ascii_punctuation_escapes(text: str) -> bool:
    """Return whether every backslash is a valid CommonMark escape."""
    cursor = 0
    while cursor < len(text):
        if text[cursor] != "\\":
            cursor += 1
            continue
        if cursor + 1 >= len(text) or text[cursor + 1] not in _ASCII_PUNCTUATION:
            return False
        cursor += 2
    return True


def _has_blank_line(text: str) -> bool:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.search(r"\n[ \t]*\n", normalized) is not None


def _label_starts_a_new_block(label: str) -> bool:
    """Detect block starts that terminate a multiline inline-label parse."""
    lines = re.split(r"\r\n|\r|\n", label)
    return any(
        _BLOCK_INTERRUPT_RE.match(line) is not None or _HTML_BLOCK_INTERRUPT_RE.match(line) is not None
        for line in lines[1:]
    )


def _normalize_reference_label(label: str) -> str:
    key = " ".join(label.split()).strip()
    return key.lower().upper()


def _reference_label_units(label: str) -> int:
    """Count parser label units, treating one backslash escape as one unit."""
    units = 0
    cursor = 0
    while cursor < len(label):
        if label[cursor] == "\\" and cursor + 1 < len(label):
            cursor += 2
        else:
            cursor += 1
        units += 1
    return units


def _strip_reference_container(line: str) -> tuple[str, int, bool]:
    """Expose a definition directly owned by quote/list block markers."""
    result = line
    quote_depth = 0
    had_list = False
    while True:
        quote = re.match(r"^ {0,3}>[ \t]?", result)
        if quote is not None:
            result = result[quote.end() :]
            quote_depth += 1
            continue
        if (
            re.fullmatch(
                r" {0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})",
                result,
            )
            is not None
        ):
            return result, quote_depth, had_list
        list_item = re.match(
            r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)])[ \t]+",
            result,
        )
        if list_item is not None:
            result = result[list_item.end() :]
            had_list = True
            continue
        return result, quote_depth, had_list


def _reference_definition_parts(line: str) -> tuple[str, str] | None:
    opening = re.match(r"^ {0,3}\[", line)
    if opening is None:
        return None
    cursor = opening.end()
    while cursor < len(line):
        if line[cursor] == "\\" and cursor + 1 < len(line) and line[cursor + 1] in _ASCII_PUNCTUATION:
            cursor += 2
            continue
        if line[cursor] == "]":
            if cursor + 1 < len(line) and line[cursor + 1] == ":":
                return line[opening.end() : cursor], line[cursor + 2 :]
            return None
        cursor += 1
    return None


def _is_valid_reference_definition_payload(
    lines: list[str],
    line_index: int,
    payload: str,
) -> bool:
    value = payload.strip(" \t")
    if not value:
        if line_index + 1 >= len(lines):
            return False
        value = lines[line_index + 1].strip(" \t")
    if not value:
        return False

    if value.startswith("<"):
        closing = value.find(">", 1)
        if closing == -1:
            return False
        destination = value[1:closing]
        if any(char in destination for char in "<>\\\x00\r\n"):
            return False
        remainder = value[closing + 1 :]
    else:
        destination_match = re.match(r"\S+", value)
        if destination_match is None:
            return False
        destination = destination_match.group(0)
        remainder = value[destination_match.end() :]

    if not remainder:
        return True
    if remainder[0] not in " \t":
        return False
    title = remainder.strip(" \t")
    return (
        bool(title)
        and "\x00" not in title
        and _TITLE_RE.fullmatch(title) is not None
        and _has_only_ascii_punctuation_escapes(title)
    )


def _reference_payload_has_inline_title(payload: str) -> bool:
    """Return whether a valid non-empty payload already carries its title."""
    value = payload.strip(" \t")
    if not value:
        return False
    if value.startswith("<"):
        closing = value.find(">", 1)
        return closing != -1 and bool(value[closing + 1 :].strip(" \t"))
    destination = re.match(r"\S+", value)
    return destination is not None and bool(value[destination.end() :].strip(" \t"))


def _is_reference_title_line(line: str) -> bool:
    title = line.strip(" \t")
    return (
        bool(title)
        and "\x00" not in title
        and _TITLE_RE.fullmatch(title) is not None
        and _has_only_ascii_punctuation_escapes(title)
    )


def _is_markdown_blank_line(line: str) -> bool:
    """Match Mistune's block blank characters, not all Unicode whitespace."""
    return re.fullmatch(r"[ \t\v\f]*", line) is not None


def _line_closes_reference_definition_block(line: str) -> bool:
    """Recognize common self-contained blocks after which a definition starts."""
    return (
        re.match(r"^ {0,3}#{1,9}(?!#)(?:[ \t]+|$)", line) is not None
        or re.fullmatch(
            r" {0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})",
            line,
        )
        is not None
    )


def _is_valid_footnote_definition_label(candidate: str) -> bool:
    key = candidate[1:]
    if not candidate.startswith("^") or not key or _reference_label_units(key) > 500:
        return False
    cursor = 0
    while cursor < len(key):
        if key[cursor] == "\\":
            if cursor + 1 >= len(key):
                return False
            cursor += 2
            continue
        if key[cursor].isspace() or key[cursor] in "[]":
            return False
        cursor += 1
    return True


def _is_valid_reference_definition_label(candidate: str) -> bool:
    if not candidate or _reference_label_units(candidate) > 500:
        return False
    cursor = 0
    while cursor < len(candidate):
        if candidate[cursor] == "\\":
            if cursor + 1 >= len(candidate):
                return False
            cursor += 2
            continue
        if candidate[cursor] in "[]":
            return False
        cursor += 1
    return True


def _has_reference_definition(text: str, label: str) -> bool:
    label_limit = 501 if label.startswith("^") else 500
    if _reference_label_units(label) > label_limit:
        return False
    normalized = _normalize_reference_label(label)
    if not normalized:
        return False
    raw_lines = re.split(r"\r\n|\r|\n", text)
    logical_lines: list[str] = []
    container_flags: list[tuple[int, bool]] = []
    for raw_line in raw_lines:
        logical, quote_depth, had_list = _strip_reference_container(raw_line)
        logical_lines.append(logical)
        container_flags.append((quote_depth, had_list))

    previous_blank = True
    previous_definition = False
    pending_continuation: str | None = None
    pending_footnote_body_line = False
    footnote_continuation_lead: int | None = None
    for line_index, line in enumerate(logical_lines):
        line_blank = _is_markdown_blank_line(line)
        if pending_footnote_body_line:
            pending_footnote_body_line = False
            if not line_blank:
                previous_blank = False
                previous_definition = True
                continue
        if footnote_continuation_lead is not None:
            continuation = re.fullmatch(
                rf" {{{footnote_continuation_lead + 1},{footnote_continuation_lead + 4}}}(?! )[^\r\n]*",
                line,
            )
            if not line_blank and continuation is not None:
                previous_blank = False
                previous_definition = True
                continue
            if not line_blank:
                footnote_continuation_lead = None

        is_continuation = False
        if pending_continuation == "destination":
            if _is_valid_reference_definition_payload(
                [line],
                0,
                line,
            ):
                is_continuation = True
                pending_continuation = None if _reference_payload_has_inline_title(line) else "title"
            else:
                pending_continuation = None
        elif pending_continuation == "title":
            if _is_reference_title_line(line):
                is_continuation = True
            pending_continuation = None

        definition = _reference_definition_parts(line)
        quote_depth, had_list = container_flags[line_index]
        quote_interrupt = quote_depth > 0 and (
            line_index == 0
            or container_flags[line_index - 1][0] < quote_depth
            or _is_markdown_blank_line(logical_lines[line_index - 1])
        )
        eligible = (
            line_index == 0
            or previous_blank
            or previous_definition
            or had_list
            or quote_interrupt
            or (line_index > 0 and _line_closes_reference_definition_block(logical_lines[line_index - 1]))
        )
        current_definition = False
        current_is_footnote = False
        if definition is not None:
            candidate, payload = definition
            footnote_syntax = _is_valid_footnote_definition_label(candidate) and (not payload or payload[0] in " \t")
            if footnote_syntax:
                # Mistune's footnotes plugin accepts an empty definition body;
                # unlike an ordinary reference definition, it can also
                # interrupt an open paragraph.  Its inline reference must
                # therefore own a following literal ``(suffix)`` instead of
                # exposing that suffix as an inline link.
                current_definition = True
                current_is_footnote = True
            elif eligible:
                # Ordinary reference definitions keep CommonMark's block
                # eligibility rules and cannot interrupt a paragraph.
                current_definition = _is_valid_reference_definition_label(
                    candidate
                ) and _is_valid_reference_definition_payload(
                    logical_lines,
                    line_index,
                    payload,
                )
            if current_definition and _normalize_reference_label(candidate) == normalized:
                return True
            if current_definition and current_is_footnote:
                pending_footnote_body_line = payload == ""
                lead = re.match(r"^( {0,3})\[", line)
                assert lead is not None
                footnote_continuation_lead = len(lead.group(1))
                pending_continuation = None
            elif current_definition:
                pending_continuation = (
                    "destination"
                    if not payload.strip(" \t")
                    else (None if _reference_payload_has_inline_title(payload) else "title")
                )
        previous_blank = line_blank
        previous_definition = current_definition or is_continuation
    return False


def _previous_label_open(text: str, close: int) -> int | None:
    """Find the escape-aware opening bracket paired with *close*."""
    depth = 1
    cursor = close - 1
    while cursor >= 0 and text[cursor] not in "\r\n":
        char = text[cursor]
        if char == "]" and not _is_backslash_escaped(text, cursor):
            depth += 1
        elif char == "[" and not _is_backslash_escaped(text, cursor):
            depth -= 1
            if depth == 0:
                return cursor
        cursor -= 1
    return None


def _is_defined_reference_tail(text: str, start: int, label: str) -> bool:
    """Detect ``[text][defined-label]`` before a literal ``(suffix)``."""
    if start <= 0 or text[start - 1] != "]" or _is_backslash_escaped(text, start - 1):
        return False
    previous_open = _previous_label_open(text, start - 1)
    if previous_open is None:
        return False
    previous_label = text[previous_open + 1 : start - 1]
    reference_label = label or previous_label
    return _has_reference_definition(text, reference_label)


def _code_span_brackets_are_balanced(text: str, start: int, end: int) -> bool:
    """Validate brackets inside a code span used as link-label content."""
    run_length = 1
    while start + run_length < end and text[start + run_length] == "`":
        run_length += 1
    content = text[start + run_length : end - run_length]
    depth = 0
    for char in content:
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth < 0:
                return False
    return depth == 0


def parse_markdown_destination(
    raw: str,
    *,
    allow_image_size: bool = False,
) -> MarkdownDestination | None:
    """Validate and split one inline-link destination expression.

    Bare destinations cannot contain unescaped whitespace.  Angle-bracket
    destinations may contain spaces, but any following content must be a
    whitespace-separated, complete title.  The Markdown image extension
    ``=WxH`` is accepted only when explicitly requested.
    """
    candidate = raw.strip(_ASCII_WHITESPACE)
    if _has_blank_line(candidate):
        return None
    if not candidate:
        return MarkdownDestination("", "")

    if candidate.startswith("<"):
        closing_angle: int | None = None
        cursor = 1
        while cursor < len(candidate):
            if candidate[cursor] == "\\":
                return None
            if candidate[cursor] in "<\x00\r\n":
                return None
            if candidate[cursor] == ">":
                closing_angle = cursor
                break
            cursor += 1
        if closing_angle is None:
            return None
        destination = candidate[1:closing_angle]
        suffix = candidate[closing_angle + 1 :]
        if suffix and suffix[0] not in _ASCII_WHITESPACE:
            return None
    else:
        cursor = 0
        while cursor < len(candidate):
            char = candidate[cursor]
            if char == "\\" and cursor + 1 < len(candidate) and candidate[cursor + 1] in _ASCII_PUNCTUATION:
                cursor += 2
                continue
            if char in _ASCII_WHITESPACE:
                break
            cursor += 1
        destination = candidate[:cursor]
        suffix = candidate[cursor:]

    if suffix:
        suffix_value = suffix.strip(_ASCII_WHITESPACE)
        valid_suffix = (
            "\x00" not in suffix_value
            and _TITLE_RE.fullmatch(suffix_value) is not None
            and _has_only_ascii_punctuation_escapes(suffix_value)
        )
        if allow_image_size:
            valid_suffix = valid_suffix or _IMAGE_SIZE_RE.fullmatch(suffix_value) is not None
        if not valid_suffix:
            return None
    return MarkdownDestination(destination, suffix)


def matching_backtick_run_end(text: str, start: int) -> int | None:
    """Return the end of an exact-length closing backtick run."""
    if start >= len(text) or text[start] != "`":
        return None
    run_length = 1
    while start + run_length < len(text) and text[start + run_length] == "`":
        run_length += 1

    cursor = start + run_length
    while cursor < len(text):
        if text[cursor] != "`":
            cursor += 1
            continue
        closing_length = 1
        while cursor + closing_length < len(text) and text[cursor + closing_length] == "`":
            closing_length += 1
        if closing_length == run_length:
            return cursor + closing_length
        cursor += closing_length
    return None


def parse_inline_link(
    text: str,
    start: int,
    *,
    image: bool,
) -> MarkdownInlineLink | None:
    """Parse a balanced, escape-aware inline link/image beginning at *start*."""
    if image:
        if not text.startswith("![", start) or _is_backslash_escaped(text, start):
            return None
        label_open = start + 1
    else:
        if start >= len(text) or text[start] != "[" or _is_backslash_escaped(text, start):
            return None
        if start > 0 and text[start - 1] == "!" and not _is_backslash_escaped(text, start - 1):
            return None
        label_open = start

    depth = 1
    cursor = label_open + 1
    label_close: int | None = None
    while cursor < len(text):
        char = text[cursor]
        if char == "\\" and cursor + 1 < len(text) and text[cursor + 1] in _ASCII_PUNCTUATION:
            cursor += 2
            continue
        if char == "`":
            code_end = matching_backtick_run_end(text, cursor)
            if code_end is not None:
                if not _code_span_brackets_are_balanced(text, cursor, code_end):
                    return None
                cursor = code_end
                continue
        if char == "[":
            depth += 1
        elif char == "]":
            backslash_run = 0
            run_cursor = cursor - 1
            while run_cursor >= 0 and text[run_cursor] == "\\":
                backslash_run += 1
                run_cursor -= 1
            if backslash_run and backslash_run % 2 == 0:
                # Mistune/CommonMark's bracket scanner consumes a positive
                # even backslash run together with the bracket as an opener.
                depth += 1
            else:
                depth -= 1
                if depth == 0:
                    label_close = cursor
                    break
        cursor += 1

    if label_close is None or label_close + 1 >= len(text) or text[label_close + 1] != "(":
        return None
    label = text[label_open + 1 : label_close]
    if _contains_active_protected_token(label) or _has_blank_line(label) or _label_starts_a_new_block(label):
        return None
    if not image and _is_defined_reference_tail(text, start, label):
        return None

    target_start = label_close + 2
    cursor = target_start
    depth = 1
    quote_char: str | None = None
    in_angle = False
    while cursor < len(text):
        char = text[cursor]
        if char == "\\" and cursor + 1 < len(text) and text[cursor + 1] in _ASCII_PUNCTUATION:
            cursor += 2
            continue
        if in_angle:
            if char == ">":
                in_angle = False
            cursor += 1
            continue
        if quote_char is not None:
            if char == quote_char:
                quote_char = None
            cursor += 1
            continue
        if char == "<":
            in_angle = True
        elif char in {'"', "'"} and depth == 1 and cursor > target_start and text[cursor - 1] in _ASCII_WHITESPACE:
            quote_char = char
        elif char == "(":
            depth += 1
        elif char == ")":
            if depth > 1 and ")" not in text[cursor + 1 :]:
                depth = 0
            else:
                depth -= 1
            if depth == 0:
                target = text[target_start:cursor]
                if _contains_active_protected_token(target):
                    return None
                if (
                    parse_markdown_destination(
                        target,
                        allow_image_size=image,
                    )
                    is None
                ):
                    return None
                return MarkdownInlineLink(
                    end=cursor + 1,
                    label=label,
                    target=target,
                    is_image=image,
                )
        cursor += 1
    return None


def encode_markdown_destination_escapes(target: str) -> str:
    """Apply CommonMark escapes while protecting URI structural literals.

    Escaped ``:`` and ``/`` still form a real URI scheme/path after Markdown
    unescaping.  Only characters whose literal value would otherwise be
    reinterpreted by URI or angle-destination parsing are percent-protected.
    """
    parts: list[str] = []
    cursor = 0
    while cursor < len(target):
        char = target[cursor]
        if char == "\\" and cursor + 1 < len(target) and target[cursor + 1] in _ASCII_PUNCTUATION:
            escaped = target[cursor + 1]
            parts.append(quote(escaped, safe="") if escaped in "#?%<>\\" else escaped)
            cursor += 2
            continue
        parts.append(char)
        cursor += 1
    return "".join(parts)


def encode_markdown_angle_destination(target: str) -> str:
    """Encode a semantic destination so a generated ``<...>`` stays valid."""
    destination = encode_markdown_destination_escapes(target)
    return destination.replace("\x00", "%00").replace("\\", "%5C").replace("<", "%3C").replace(">", "%3E")


def escape_unescaped_pipes(text: str) -> str:
    """Escape table-delimiting pipes while preserving escape parity."""
    parts: list[str] = []
    backslash_run = 0
    for char in text:
        if char == "\\":
            parts.append(char)
            backslash_run += 1
            continue
        if char == "|" and backslash_run % 2 == 0:
            parts.append("\\")
        parts.append(char)
        backslash_run = 0
    return "".join(parts)


def escape_markdown_label(label: str) -> str:
    """Escape generated link-label delimiters while preserving visible text."""
    return label.replace("\\", r"\\").replace("[", r"\[").replace("]", r"\]")


def escape_markdown_source_literal(source: str) -> str:
    """Make Markdown syntax render as its exact source characters."""
    return "".join(f"\\{char}" if char in _ASCII_PUNCTUATION else char for char in source)
