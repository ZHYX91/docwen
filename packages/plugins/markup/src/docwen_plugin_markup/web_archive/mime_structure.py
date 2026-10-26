"""Select the related root and keep each part's own RFC 2557 URI context."""

from __future__ import annotations

from dataclasses import dataclass
from email import policy
from email.message import Message
from email.parser import BytesParser

from docwen_plugin_markup.web_archive.resource_identity import resolve_mhtml_uri


@dataclass(frozen=True)
class MhtmlPart:
    message: Message
    path: tuple[int, ...]
    location: str
    base_uri: str
    related_ancestors: tuple[tuple[int, ...], ...]


def _cid(value: object) -> str:
    return str(value or "").strip().strip("<>")


def _children(part: Message) -> list[Message]:
    payload = part.get_payload()
    if not isinstance(payload, list):
        return []
    children: list[Message] = []
    for child in payload:
        if not isinstance(child, Message):
            raise ValueError("MHTML multipart body contains an unsupported part")
        children.append(child)
    return children


def _html_root(part: Message) -> Message:
    kind = part.get_content_type().lower()
    if kind == "text/html" and not part.is_multipart():
        return part
    children = _children(part)
    if kind == "multipart/related":
        if not children:
            raise ValueError("MHTML related root has no body parts")
        start = part.get_param("start")
        if start is None:
            root = children[0]
        else:
            matches = [child for child in children if _cid(child.get("Content-ID")) == _cid(start)]
            if not _cid(start) or len(matches) != 1:
                raise ValueError("MHTML declared related root is missing or ambiguous")
            root = matches[0]
        declared_type = part.get_param("type")
        if declared_type and str(declared_type).lower() != root.get_content_type().lower():
            raise ValueError("MHTML related root does not match its declared media type")
        return _html_root(root)
    if kind == "multipart/alternative":
        # RFC 2046 orders alternatives by increasing preference. Select the
        # last supported representation, never an unrelated HTML sibling.
        for child in reversed(children):
            if child.get_content_type().lower() in {"text/html", "multipart/related", "multipart/alternative"}:
                return _html_root(child)
        raise ValueError("MHTML root alternative has no supported HTML representation")
    raise ValueError("MHTML selected root has no usable text/html body (unsupported representation)")


def parse_mhtml_structure(raw: bytes) -> tuple[MhtmlPart, list[MhtmlPart]]:
    """Return the chosen HTML and visible resource contexts, without flattening bases."""
    message = BytesParser(policy=policy.default).parsebytes(raw)
    root = _html_root(message)
    entries: list[MhtmlPart] = []

    def visit(part: Message, path: tuple[int, ...], parent_base: str, scopes: tuple[tuple[int, ...], ...]) -> None:
        raw_location = str(part.get("Content-Location") or "").strip()
        # cid URLs match Content-ID only (RFC 2557 section 8.3).
        location = (
            resolve_mhtml_uri(parent_base, raw_location)
            if raw_location and not raw_location.lower().startswith("cid:")
            else ""
        )
        base = location or parent_base
        if not location and part.get("Content-Base"):
            # Conservative read-only support permitted by RFC 2557 section 12.
            base = resolve_mhtml_uri(parent_base, str(part.get("Content-Base")))
        if part.get_content_type().lower() == "multipart/related":
            scopes = (*scopes, path)
        entries.append(MhtmlPart(part, path, location, base, scopes))
        for index, child in enumerate(_children(part)):
            visit(child, (*path, index), base, scopes)

    visit(message, (), "thismessage:/", ())
    selected = next(entry for entry in entries if entry.message is root)
    visible = [
        entry
        for entry in entries
        if selected.related_ancestors[: len(entry.related_ancestors)] == entry.related_ancestors
    ]
    return selected, visible
