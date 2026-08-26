"""Target-aware processing for non-embed Markdown and wiki links."""

from __future__ import annotations

import logging
import re
import secrets
from collections.abc import Callable, Sequence
from pathlib import Path, PureWindowsPath
from urllib.parse import quote, unquote, urlsplit

from docwen_core.links._embed_dispatch import _is_table_context
from docwen_core.links._error_semantics import LinkErrorKind, dispatch_error_output
from docwen_core.links._markdown_inline import (
    _ACTIVE_PROTECTED_TOKENS,
    _contains_active_protected_token,
    _is_backslash_escaped,
    encode_markdown_angle_destination,
    encode_markdown_destination_escapes,
    escape_markdown_label,
    escape_markdown_source_literal,
    escape_unescaped_pipes,
    matching_backtick_run_end,
    parse_inline_link,
    parse_markdown_destination,
)
from docwen_core.links._patterns import WIKI_EMBED_PATTERN
from docwen_core.links._resolver import resolve_file_path

logger = logging.getLogger(__name__)

_WIKI_NON_EMBED_RE = re.compile(r"(?<!!)" + WIKI_EMBED_PATTERN[1:])
_WIKI_EMBED_RE = re.compile(WIKI_EMBED_PATTERN)
_ANGLE_AUTOLINK_RE = re.compile(
    r"<[A-Za-z][A-Za-z0-9.+-]{1,31}:[^<>\x00-\x20]*>"
    r"|<[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9]"
    r"(?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*>"
)
_INLINE_HTML_RE = re.compile(
    r"<[A-Za-z][A-Za-z0-9-]*"
    r"(?:\s+[A-Za-z_:][A-Za-z0-9_.:-]*"
    r"(?:\s*=\s*(?:[^ !\"'=<>`]+|'[^']*?'|\"[^\"]*?\"))?)*\s*/?>"
    r"|</[A-Za-z][A-Za-z0-9-]*\s*>"
    r"|<!--(?!>|->)(?:(?!--)[\s\S])+?(?<!-)-->"
    r"|<\?[\s\S]+?\?>"
    r"|<![A-Z][\s\S]+?>"
    r"|<!\[CDATA[\s\S]+?\]\]>"
)
_INLINE_MATH_RE = re.compile(r"\$(?!\s)(?:[^$\\]|\\.)+?(?!\s)\$")
_HTML_PRE_TAGS = frozenset({"pre", "script", "style", "textarea"})
_HTML_BLOCK_TAGS = frozenset(
    {
        "address",
        "article",
        "aside",
        "base",
        "basefont",
        "blockquote",
        "body",
        "caption",
        "center",
        "col",
        "colgroup",
        "dd",
        "details",
        "dialog",
        "dir",
        "div",
        "dl",
        "dt",
        "fieldset",
        "figcaption",
        "figure",
        "footer",
        "form",
        "frame",
        "frameset",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "head",
        "header",
        "hr",
        "html",
        "iframe",
        "legend",
        "li",
        "link",
        "main",
        "menu",
        "menuitem",
        "meta",
        "nav",
        "noframes",
        "ol",
        "optgroup",
        "option",
        "p",
        "param",
        "section",
        "source",
        "summary",
        "table",
        "tbody",
        "td",
        "tfoot",
        "th",
        "thead",
        "title",
        "tr",
        "track",
        "ul",
    }
)


def _line_completes_markdown_block(
    line: str,
    *,
    setext_context: bool = False,
) -> bool:
    """Return whether *line* permits an immediately following indented block."""
    without_eol = line.rstrip("\r\n")
    if re.match(r"^ {0,3}#{1,9}(?!#)(?:[ \t]+|$)", without_eol):
        return True
    if setext_context and re.match(r"^ {0,3}(?:=+|-+)[ \t]*$", without_eol):
        return True
    return bool(
        re.match(
            r"^ {0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})$",
            without_eol,
        )
    )


def _line_can_precede_setext_heading(line: str) -> bool:
    """Conservatively recognize paragraph text eligible for setext syntax."""
    without_eol = line.rstrip("\r\n")
    if not without_eol.strip():
        return False
    if re.match(r"^ {0,3}(?:#{1,9}(?!#)(?:[ \t]+|$)|>|`{3,}|~{3,})", without_eol):
        return False
    if re.match(r"^ {0,3}(?:[-+*](?:[ \t]+|$)|\d+[.)](?:[ \t]+|$))", without_eol):
        return False
    return not _line_completes_markdown_block(without_eol)


def _unescape_pipe(text: str) -> str:
    """Unescape a pipe in a wiki-link target or display value."""
    return text.replace("\\|", "|")


def _trim_bare_url_candidate(candidate: str) -> tuple[str, str]:
    """Separate a conservative bare URL from trailing prose punctuation."""
    trailing = ""
    while candidate and candidate[-1] in ".,;:!?":
        trailing = candidate[-1] + trailing
        candidate = candidate[:-1]
    while candidate.endswith(")") and candidate.count(")") > candidate.count("("):
        trailing = ")" + trailing
        candidate = candidate[:-1]
    while candidate.endswith("]") and candidate.count("]") > candidate.count("["):
        trailing = "]" + trailing
        candidate = candidate[:-1]
    while candidate.endswith("}") and candidate.count("}") > candidate.count("{"):
        trailing = "}" + trailing
        candidate = candidate[:-1]
    return candidate, trailing


def _bare_url_end(text: str, start: int) -> int | None:
    """Return the visible URL end for the opt-in bare-URL grammar."""
    if not (text.startswith("https://", start) or text.startswith("http://", start)):
        return None
    if start > 0 and (text[start - 1].isalnum() or text[start - 1] in "_/"):
        return None
    end = start
    while end < len(text) and not text[end].isspace() and text[end] not in "<>":
        end += 1
    url, trailing = _trim_bare_url_candidate(text[start:end])
    if not url:
        return None
    try:
        hostname = urlsplit(url).hostname or ""
    except ValueError:
        return None
    if not hostname or not any(char.isalnum() for char in hostname):
        return None
    return end - len(trailing)


def _protected_inline_atom_end(
    text: str,
    start: int,
    *,
    protect_bare_urls: bool,
) -> int | None:
    """Return the end of a renderer atom that link policy must not enter."""
    if text[start] == "<":
        match = _ANGLE_AUTOLINK_RE.match(text, start)
        if match is None:
            match = _INLINE_HTML_RE.match(text, start)
        if match is not None:
            return match.end()
    if text[start] == "$" and not _is_backslash_escaped(text, start):
        match = _INLINE_MATH_RE.match(text, start)
        if match is not None:
            return match.end()
    if protect_bare_urls and text[start] == "h":
        return _bare_url_end(text, start)
    return None


def _is_indented_code_line(line: str) -> bool:
    """Recognize a four-column CommonMark indentation prefix."""
    columns = 0
    for char in line:
        if char == " ":
            columns += 1
        elif char == "\t":
            columns += 4 - (columns % 4)
        else:
            break
        if columns >= 4:
            return True
    return False


def _html_block_mode(
    line: str,
    *,
    allow_type7: bool,
) -> tuple[str | None, bool] | None:
    """Return ``(end marker, until blank)`` for a Mistune HTML block start."""
    leading = re.match(r"^ {0,3}", line)
    assert leading is not None
    content = line[leading.end() :]
    if content.startswith("<!--"):
        return "-->", False
    if content.startswith("<?"):
        return "?>", False
    if content.startswith("<![CDATA["):
        return "]]>", False
    if re.match(r"<![A-Z]", content):
        return ">", False

    match = re.match(
        r"</?([A-Za-z][A-Za-z0-9-]*)\b",
        content,
        re.IGNORECASE,
    )
    if match is None:
        return None
    marker = match.group(0).lower()
    tag = (match.group(1) or "").lower()
    if tag in _HTML_PRE_TAGS and not marker.startswith("</"):
        return f"</{tag}>", False
    if tag in _HTML_BLOCK_TAGS or tag in _HTML_PRE_TAGS:
        return None, True
    candidate = line.lstrip(" ").rstrip(" \t\r\n")
    generic_self_closing = candidate.startswith("<") and not candidate.startswith("</") and candidate.endswith("/>")
    if allow_type7 and not generic_self_closing and _INLINE_HTML_RE.fullmatch(candidate) is not None:
        return None, True
    return None


def _strip_block_quote_prefix(line: str) -> str:
    """Strip repeated CommonMark blockquote markers for block detection."""
    result = line
    while True:
        match = re.match(r"^ {0,3}>[ \t]?", result)
        if match is None:
            return result
        result = result[match.end() :]


def _container_block_line(line: str) -> tuple[str, bool]:
    """Return parser-visible block content and whether a list marker owned it."""
    quoted = _strip_block_quote_prefix(line)
    match = re.match(r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)])[ \t]+(.*)$", quoted)
    if match is None:
        return quoted, False
    return match.group(1), True


def _split_fenced_code_blocks(md: str) -> list[tuple[str, bool]]:
    """Split Markdown into ``(segment, is_fenced_code)`` pairs."""
    segments: list[tuple[str, bool]] = []
    if not md:
        return segments

    lines = md.splitlines(keepends=True)
    buffer: list[str] = []
    in_fence = False
    in_math_block = False
    html_end_marker: str | None = None
    html_until_blank = False
    fence = ""
    fence_in_list = False
    fence_in_quote = False
    in_indented_code = False
    indented_code_allowed = True
    previous_line_can_be_setext = False

    for line_index, line in enumerate(lines):
        line_has_quote_marker = re.match(r"^ {0,3}>", line) is not None
        quoted_line = _strip_block_quote_prefix(line)
        block_line, line_has_list_marker = _container_block_line(line)
        if html_end_marker is not None or html_until_blank:
            if html_until_blank and not quoted_line.strip():
                segments.append(("".join(buffer), True))
                buffer = []
                html_until_blank = False
            else:
                buffer.append(line)
                if html_end_marker is not None and html_end_marker in line.lower():
                    segments.append(("".join(buffer), True))
                    buffer = []
                    html_end_marker = None
                    indented_code_allowed = True
                    previous_line_can_be_setext = False
                continue
        if in_math_block:
            buffer.append(line)
            if re.match(r"^\$\$[ \t]*(?:\r?\n)?$", line):
                segments.append(("".join(buffer), True))
                buffer = []
                in_math_block = False
                indented_code_allowed = True
                previous_line_can_be_setext = False
            continue
        if not in_fence:
            html_mode = _html_block_mode(
                block_line,
                allow_type7=indented_code_allowed,
            )
            if html_mode is not None:
                if buffer:
                    segments.append(("".join(buffer), False))
                    buffer = []
                html_end_marker, html_until_blank = html_mode
                buffer.append(line)
                if html_end_marker is not None and html_end_marker in line.lower():
                    segments.append(("".join(buffer), True))
                    buffer = []
                    html_end_marker = None
                in_indented_code = False
                indented_code_allowed = False
                previous_line_can_be_setext = False
                continue
            single_line_math = re.match(
                r"^ {0,3}\$\$[^\r\n]+?\$\$[ \t]*(?:\r?\n)?$",
                line,
            )
            if single_line_math:
                if buffer:
                    segments.append(("".join(buffer), False))
                    buffer = []
                segments.append((line, True))
                in_indented_code = False
                indented_code_allowed = True
                previous_line_can_be_setext = False
                continue
            if re.match(r"^ {0,3}\$\$[ \t]*(?:\r?\n)$", line) and any(
                re.match(r"^\$\$[ \t]*(?:\r?\n)?$", candidate) is not None for candidate in lines[line_index + 1 :]
            ):
                if buffer:
                    segments.append(("".join(buffer), False))
                    buffer = []
                in_math_block = True
                buffer.append(line)
                in_indented_code = False
                indented_code_allowed = False
                previous_line_can_be_setext = False
                continue
            match = re.match(r"^ {0,3}(`{3,}|~{3,})([^\r\n]*)", block_line)
            if match and (match.group(1)[0] == "~" or "`" not in match.group(2)):
                if buffer:
                    segments.append(("".join(buffer), False))
                    buffer = []
                in_fence = True
                fence = match.group(1)
                fence_in_list = line_has_list_marker
                fence_in_quote = line_has_quote_marker
                buffer.append(line)
                in_indented_code = False
                indented_code_allowed = False
                previous_line_can_be_setext = False
                continue
            if _is_indented_code_line(block_line) and (in_indented_code or indented_code_allowed):
                if buffer:
                    segments.append(("".join(buffer), False))
                    buffer = []
                segments.append((line, True))
                in_indented_code = True
                indented_code_allowed = True
                previous_line_can_be_setext = False
                continue
            if line.strip():
                in_indented_code = False
            buffer.append(line)
            completes_block = _line_completes_markdown_block(
                line,
                setext_context=previous_line_can_be_setext,
            )
            indented_code_allowed = not line.strip() or completes_block
            previous_line_can_be_setext = not completes_block and _line_can_precede_setext_heading(line)
            continue

        if fence_in_quote and line.strip() and not line_has_quote_marker:
            segments.append(("".join(buffer), True))
            buffer = [line]
            in_fence = False
            fence = ""
            fence_in_list = False
            fence_in_quote = False
            in_indented_code = False
            completes_block = _line_completes_markdown_block(
                line,
                setext_context=previous_line_can_be_setext,
            )
            indented_code_allowed = not line.strip() or completes_block
            previous_line_can_be_setext = not completes_block and _line_can_precede_setext_heading(line)
            continue

        buffer.append(line)
        closing_line = block_line.lstrip(" \t") if fence_in_list else block_line
        closing = re.match(
            r"^ {0,3}(`{3,}|~{3,})[ \t]*(?:\r?\n)?$",
            closing_line,
        )
        if closing and closing.group(1)[0] == fence[0] and len(closing.group(1)) >= len(fence):
            segments.append(("".join(buffer), True))
            buffer = []
            in_fence = False
            fence = ""
            fence_in_list = False
            fence_in_quote = False
            # A fenced block is a complete block boundary.  A following
            # four-space line may therefore begin indented code even without
            # an intervening blank line; only paragraph continuations are
            # forbidden from doing so.
            indented_code_allowed = True
            previous_line_can_be_setext = False

    if buffer:
        segments.append(
            (
                "".join(buffer),
                in_fence or in_math_block or html_end_marker is not None or html_until_blank,
            )
        )
    return segments


def split_markdown_block_segments(md: str) -> list[tuple[str, bool]]:
    """Split Markdown into visible and renderer-protected block segments."""

    return _split_fenced_code_blocks(md)


def _split_inline_code_spans(
    segment: str,
    *,
    protect_bare_urls: bool = True,
) -> list[tuple[str, bool]]:
    """Split non-fenced Markdown into visible text and protected atoms."""
    parts: list[tuple[str, bool]] = []
    if not segment:
        return parts

    cursor = 0
    index = 0
    while index < len(segment):
        if not protect_bare_urls and segment[index] == "h":
            bare_url_end = _bare_url_end(segment, index)
            if bare_url_end is not None:
                index = bare_url_end
                continue
        construct = parse_inline_link(segment, index, image=True)
        if construct is None:
            construct = parse_inline_link(segment, index, image=False)
        if construct is not None:
            index = construct.end
            continue
        wiki_match = _WIKI_EMBED_RE.match(segment, index)
        if wiki_match is None:
            wiki_match = _WIKI_NON_EMBED_RE.match(segment, index)
        if wiki_match is not None:
            index = wiki_match.end()
            continue
        atom_end = _protected_inline_atom_end(
            segment,
            index,
            protect_bare_urls=protect_bare_urls,
        )
        if atom_end is not None:
            if cursor < index:
                parts.append((segment[cursor:index], False))
            parts.append((segment[index:atom_end], True))
            cursor = atom_end
            index = atom_end
            continue
        if segment[index] != "`":
            index += 1
            continue
        if _is_backslash_escaped(segment, index):
            index += 1
            continue

        code_end = matching_backtick_run_end(segment, index)
        if code_end is None:
            run_length = 1
            while index + run_length < len(segment) and segment[index + run_length] == "`":
                run_length += 1
            index += run_length
            continue

        fence_length = 1
        while index + fence_length < len(segment) and segment[index + fence_length] == "`":
            fence_length += 1
        if code_end <= index + fence_length:
            index += fence_length
            continue

        if cursor < index:
            parts.append((segment[cursor:index], False))
        parts.append((segment[index:code_end], True))
        cursor = code_end
        index = code_end

    if cursor < len(segment):
        parts.append((segment[cursor:], False))
    return parts or [(segment, False)]


def split_markdown_inline_segments(
    segment: str,
    *,
    protect_bare_urls: bool = True,
) -> list[tuple[str, bool]]:
    """Split one block segment into visible and protected inline atoms."""

    return _split_inline_code_spans(
        segment,
        protect_bare_urls=protect_bare_urls,
    )


def _map_visible_markdown(
    text: str,
    transform: Callable[[str], str],
    *,
    protect_bare_urls: bool = True,
) -> str:
    """Apply *transform* once with renderer atoms replaced by scoped tokens."""
    protected: dict[str, str] = {}
    active_tokens: set[str] = set()
    nonce = secrets.token_hex(16)

    def mask(value: str) -> str:
        index = len(protected)
        token = f"<DOCWEN-VISIBLE-{nonce}-{index}>"
        while token in text or token in protected:
            index += 1
            token = f"<DOCWEN-VISIBLE-{nonce}-{index}>"
        leading_match = re.match(r"[ \t]*", value)
        assert leading_match is not None
        leading = leading_match.group(0)
        line_endings = "".join(re.findall(r"\r\n|\r|\n", value))
        # Retain horizontal indentation so a protected indented-code line can
        # still carry surrounding block context.  In particular, Mistune's
        # footnote grammar owns a 1..4-space continuation even when a following
        # tab makes that physical line look like four-column indented code.
        placeholder = leading + token + line_endings
        protected[placeholder] = value
        active_tokens.add(token)
        return placeholder

    block_masked = "".join(
        mask(segment) if is_protected else segment for segment, is_protected in _split_fenced_code_blocks(text)
    )
    inline_masked = "".join(
        mask(segment) if is_protected else segment
        for segment, is_protected in _split_inline_code_spans(
            block_masked,
            protect_bare_urls=protect_bare_urls,
        )
    )
    context_token = _ACTIVE_PROTECTED_TOKENS.set(frozenset(active_tokens))
    try:
        result = transform(inline_masked)
    finally:
        _ACTIVE_PROTECTED_TOKENS.reset(context_token)
    for placeholder, original in reversed(protected.items()):
        result = result.replace(placeholder, original)
    return result


def _replace_markdown_links(
    segment: str,
    replacer: Callable[[str, str, str], str | None],
    *,
    table_safe: bool = False,
) -> str:
    """Replace standard Markdown links while supporting nested URL parens."""
    if not segment:
        return segment

    parts: list[str] = []
    cursor = 0
    index = 0
    while index < len(segment):
        image_construct = parse_inline_link(segment, index, image=True)
        if image_construct is not None:
            index = image_construct.end
            continue
        if segment[index] != "[":
            index += 1
            continue

        construct = parse_inline_link(segment, index, image=False)
        if construct is None:
            index += 1
            continue
        original = segment[index : construct.end]
        replacement = replacer(construct.label, construct.target, original)
        in_table = table_safe and _is_table_context(segment, index)
        if replacement is None and in_table:
            replacement = original
        if replacement is not None:
            if in_table:
                replacement = escape_unescaped_pipes(replacement)
            parts.append(segment[cursor:index])
            parts.append(replacement)
            cursor = construct.end
        index = construct.end

    parts.append(segment[cursor:])
    return "".join(parts)


def _canonical_local_docx_target(
    raw_target: str,
    source_file_path: str,
) -> str | None:
    """Return an absolute file-URI Markdown target for a local destination."""
    parsed_target = parse_markdown_destination(raw_target)
    if parsed_target is None or not parsed_target.destination:
        return None
    destination = encode_markdown_destination_escapes(parsed_target.destination)
    suffix = parsed_target.suffix
    if not destination or destination.startswith("#"):
        return None

    if re.match(r"^[A-Za-z]:[/\\]", destination):
        path_and_query, separator, fragment = destination.partition("#")
        path_text, query_separator, query = path_and_query.partition("?")
        target = PureWindowsPath(unquote(path_text)).as_uri()
        if query_separator:
            target += f"?{query}"
        if separator:
            target += f"#{quote(unquote(fragment), safe='/-._~')}"
        return f"<{target}>{suffix}"

    parsed = urlsplit(destination)
    if parsed.scheme or parsed.netloc or not parsed.path:
        return None

    local_path = Path(unquote(parsed.path))
    if not local_path.is_absolute():
        local_path = Path(source_file_path).parent / local_path
    target = local_path.resolve().as_uri()
    if parsed.query:
        target += f"?{parsed.query}"
    if parsed.fragment:
        target += f"#{quote(unquote(parsed.fragment), safe='/-._~')}"
    return f"<{target}>{suffix}"


def _process_non_embed_links(
    text: str,
    *,
    source_file_path: str,
    wiki_mode: str = "keep",
    markdown_mode: str = "keep",
    search_dirs: Sequence[str] | None = None,
    target_format: str | None = None,
    on_not_found: str = "placeholder",
    canonicalize_local_docx_targets: bool = False,
    table_safe: bool = False,
) -> str:
    """Process links outside fenced and inline code.

    ``hyperlink`` becomes renderer-ready Markdown for DOCX.  XLSX and CSV
    deliberately retain the source syntax because they have no equivalent
    renderer boundary.  A DOCX ``keep`` policy escapes standard Markdown link
    syntax so the renderer emits it literally.
    """
    if not text:
        return text

    normalized_target = (target_format or "").strip().lower()

    def _replace_wiki(match: re.Match[str]) -> str:
        target = _unescape_pipe((match.group(1) or "").strip())
        display_raw = match.group(2)
        display = _unescape_pipe(display_raw.strip()) if display_raw else target

        if wiki_mode == "keep" and normalized_target == "docx":
            return escape_markdown_source_literal(match.group(0))
        if wiki_mode == "extract_text":
            return display or target
        if wiki_mode == "remove":
            return ""

        should_resolve = wiki_mode == "resolve" or (wiki_mode == "hyperlink" and normalized_target == "docx")
        if not should_resolve:
            return match.group(0)
        if not target:
            return display

        encoded_target = encode_markdown_destination_escapes(target)
        path_and_query, separator, raw_fragment = encoded_target.partition("#")
        raw_path, query_separator, raw_query = path_and_query.partition("?")
        target_path = unquote(raw_path)
        fragment = unquote(raw_fragment)
        if not raw_path:
            # Word output has no heading-bookmark contract yet.  Rendering a
            # fragment-only target as a hyperlink would therefore create only
            # underlined text with no navigation semantics.  Match both
            # reference implementations and degrade honestly to display text.
            return display

        if re.match(r"^(?:https?|ftp|file|mailto):", raw_path, re.IGNORECASE):
            destination = encode_markdown_angle_destination(target)
            return f"[{escape_markdown_label(display)}](<{destination}>)"
        if encoded_target.startswith("//"):
            return display

        resolved = resolve_file_path(
            raw_path,
            source_file_path,
            search_dirs=list(search_dirs) if search_dirs is not None else None,
        )
        if resolved is None:
            heading = fragment if separator and not fragment.startswith("^") else None
            block_id = fragment[1:] if separator and fragment.startswith("^") else None
            error_output = dispatch_error_output(
                LinkErrorKind.FILE_NOT_FOUND,
                on_not_found,
                target_path,
                heading=heading,
                block_id=block_id,
                original_link=match.group(0),
            )
            if normalized_target == "docx" and error_output == match.group(0):
                return escape_markdown_source_literal(error_output)
            return error_output
        if canonicalize_local_docx_targets and normalized_target == "docx":
            path_text = Path(resolved).resolve().as_uri()
            if query_separator:
                path_text = f"{path_text}?{raw_query}"
            if separator:
                path_text = f"{path_text}#{quote(unquote(fragment), safe='/-._~')}"
            return f"[{escape_markdown_label(display)}](<{path_text}>)"

        try:
            source_dir = Path(source_file_path).parent
            path_text = Path(resolved).relative_to(source_dir).as_posix()
        except ValueError:
            path_text = resolved.replace("\\", "/")
        path_text = quote(path_text, safe="/:._-~")
        if query_separator:
            path_text = f"{path_text}?{raw_query}"
        if separator:
            path_text = f"{path_text}#{quote(fragment, safe='/-._~')}"
        return f"[{escape_markdown_label(display)}](<{path_text}>)"

    def _replace_markdown(display: str, target: str, original: str) -> str | None:
        display_text = display.strip()
        target_text = target.strip(" \t\r\n")
        if markdown_mode == "extract_text":
            return display_text or target_text or ""
        if markdown_mode == "remove":
            return ""
        if markdown_mode == "keep" and normalized_target == "docx":
            return escape_markdown_source_literal(original)
        if markdown_mode == "hyperlink" and normalized_target == "docx":
            parsed_target = parse_markdown_destination(target_text)
            if parsed_target is None:
                return None
            if parsed_target.destination.startswith("#"):
                return display
            canonical_target = None
            if canonicalize_local_docx_targets:
                canonical_target = _canonical_local_docx_target(
                    target_text,
                    source_file_path,
                )
            if canonical_target is None:
                destination = encode_markdown_angle_destination(parsed_target.destination)
                canonical_target = f"<{destination}>{parsed_target.suffix}"
            # ``display`` is the original, validated Markdown label source.
            # Re-encoding its backslashes would change escape/formatting
            # semantics (for example ``\*literal\*``).
            return f"[{display}]({canonical_target})"
        return None

    def _replace_wiki_links(segment: str) -> str:
        parts: list[str] = []
        cursor = 0
        index = 0
        while index < len(segment):
            construct = parse_inline_link(segment, index, image=True)
            if construct is None:
                construct = parse_inline_link(segment, index, image=False)
            if construct is not None:
                index = construct.end
                continue
            match = _WIKI_NON_EMBED_RE.match(segment, index)
            if match is None:
                index += 1
                continue
            if _contains_active_protected_token(match.group(0)):
                index = match.end()
                continue
            parts.append(segment[cursor:index])
            replacement = _replace_wiki(match)
            if table_safe and _is_table_context(segment, index):
                replacement = escape_unescaped_pipes(replacement)
            parts.append(replacement)
            index = match.end()
            cursor = index
        parts.append(segment[cursor:])
        return "".join(parts)

    def _process_visible(segment: str) -> str:
        result = segment
        # Apply the standard-Markdown policy first.  A DOCX wiki hyperlink is
        # generated as Markdown below and must not then be mistaken for an
        # input link governed by ``markdown_mode`` (notably ``keep``).
        should_process_markdown = (
            table_safe
            or markdown_mode in ("extract_text", "remove")
            or (normalized_target == "docx" and markdown_mode in ("keep", "hyperlink"))
        )
        if should_process_markdown:
            result = _replace_markdown_links(
                result,
                _replace_markdown,
                table_safe=table_safe,
            )
        if wiki_mode in ("keep", "extract_text", "remove", "resolve", "hyperlink"):
            result = _replace_wiki_links(result)
        return result

    result = _map_visible_markdown(text, _process_visible)
    logger.debug("Processed non-embed links: input=%d output=%d", len(text), len(result))
    return result
