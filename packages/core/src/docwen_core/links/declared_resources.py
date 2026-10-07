"""Fail-closed resolution inside a request-declared virtual input root."""

from __future__ import annotations

import hashlib
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import PurePosixPath
from urllib.parse import unquote, urlsplit

from docwen_core.links._non_embed import _map_visible_markdown

MARKDOWN_RESOURCE_BINDINGS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["authored_sha256", "images"],
    "properties": {
        "authored_sha256": {"type": "string", "pattern": "^[0-9a-f]{64}$"},
        "images": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["authored_token", "logical_path"],
                "properties": {
                    "authored_token": {"type": "string", "minLength": 1},
                    "logical_path": {"type": "string", "minLength": 1},
                },
            },
        },
        "wiki_links": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["authored_token", "href"],
                "properties": {
                    "authored_token": {"type": "string", "minLength": 1},
                    "href": {"type": "string", "minLength": 1},
                },
            },
        },
    },
}

_WIKI_LINK_RE = re.compile(r"!?\[\[(?:[^\]\\]|\\.)+\]\]")
_REMOTE_LINK_SCHEMES = frozenset({"ftp", "http", "https", "mailto"})
_DECLARED_WIKI_SCHEMES = frozenset({"http", "https", "mailto", "obsidian"})


class DeclaredResourceError(ValueError):
    """Raised when a local resource is not a declared request input."""


@dataclass(frozen=True, slots=True)
class DeclaredResourceResolver:
    """Resolve logical Markdown targets without consulting the physical filesystem."""

    source_logical_path: str
    resources: dict[str, str]
    bindings: dict[str, str] = field(default_factory=dict)
    wiki_links: dict[str, str] = field(default_factory=dict)

    def with_bindings(self, source: str, value: object) -> DeclaredResourceResolver:
        """Authenticate token bindings against this source and the declared inventory."""
        if value is None:
            return self
        if not isinstance(value, dict):
            raise DeclaredResourceError("invalid Markdown resource bindings")
        allowed_keys = {"authored_sha256", "images", "wiki_links"}
        if set(value) - allowed_keys or not {"authored_sha256", "images"} <= set(value):
            raise DeclaredResourceError("invalid Markdown resource bindings")
        if value["authored_sha256"] != hashlib.sha256(source.encode("utf-8")).hexdigest():
            raise DeclaredResourceError("Markdown resource bindings source hash mismatch")
        images = value["images"]
        if not isinstance(images, list):
            raise DeclaredResourceError("invalid Markdown image bindings")
        visible_tokens: set[str] = set()
        visible_wiki_links: set[str] = set()

        def collect(segment: str) -> str:
            from docwen_core.links._markdown_inline import parse_inline_link
            from docwen_core.links._patterns import WIKI_EMBED_PATTERN

            index = 0
            while index < len(segment):
                construct = parse_inline_link(segment, index, image=True)
                if construct is not None:
                    visible_tokens.add(segment[index : construct.end])
                    index = construct.end
                    continue
                link = parse_inline_link(segment, index, image=False)
                if link is not None:
                    index = link.end
                    continue
                wiki = re.compile(WIKI_EMBED_PATTERN).match(segment, index)
                if wiki is not None:
                    visible_tokens.add(wiki.group(0))
                    index = wiki.end()
                    continue
                navigation = _WIKI_LINK_RE.match(segment, index)
                if navigation is not None and not navigation.group(0).startswith("!"):
                    visible_wiki_links.add(navigation.group(0))
                    index = navigation.end()
                    continue
                index += 1
            return segment

        _map_visible_markdown(source, collect)
        bindings: dict[str, str] = {}
        for item in images:
            if not isinstance(item, dict) or set(item) != {"authored_token", "logical_path"}:
                raise DeclaredResourceError("invalid Markdown image binding")
            token, logical = item["authored_token"], item["logical_path"]
            if not isinstance(token, str) or token not in visible_tokens:
                raise DeclaredResourceError("image binding is not a visible authored image")
            if not isinstance(logical, str) or logical not in self.resources:
                raise DeclaredResourceError("image binding targets an undeclared resource")
            if token in bindings and bindings[token] != logical:
                raise DeclaredResourceError("conflicting image bindings")
            bindings[token] = logical
        wiki_links: dict[str, str] = {}
        raw_wiki_links = value.get("wiki_links", [])
        if not isinstance(raw_wiki_links, list):
            raise DeclaredResourceError("invalid Markdown wiki link bindings")
        for item in raw_wiki_links:
            if not isinstance(item, dict) or set(item) != {"authored_token", "href"}:
                raise DeclaredResourceError("invalid Markdown wiki link binding")
            token, href = item["authored_token"], item["href"]
            if not isinstance(token, str) or token not in visible_wiki_links:
                raise DeclaredResourceError("wiki link binding is not a visible authored wiki link")
            if not isinstance(href, str):
                raise DeclaredResourceError("invalid Markdown wiki link href")
            parsed_href = urlsplit(href)
            if parsed_href.scheme.lower() not in _DECLARED_WIKI_SCHEMES:
                raise DeclaredResourceError("unsupported declared wiki link target")
            if token in wiki_links and wiki_links[token] != href:
                raise DeclaredResourceError("conflicting wiki link bindings")
            wiki_links[token] = href
        return DeclaredResourceResolver(
            self.source_logical_path,
            self.resources,
            bindings,
            wiki_links,
        )

    def resolve_image(self, authored_token: str, target: str) -> str:
        logical = self.bindings.get(authored_token)
        return self.resources[logical] if logical is not None else self.resolve(target)

    def resolve_wiki_link(self, authored_token: str, target: str) -> str | None:
        """Return the authenticated navigation URI for one authored WikiLink."""

        del target
        return self.wiki_links.get(authored_token)

    def resolve(self, target: str) -> str:
        decoded_target = unquote(target)
        normalized_physical = decoded_target.replace("/", "\\")
        for physical in self.resources.values():
            if normalized_physical == physical.replace("/", "\\"):
                return physical
        parsed = urlsplit(target)
        if parsed.scheme or parsed.netloc or target.startswith(("/", "\\")) or "\\" in target:
            raise DeclaredResourceError("resource target must be a relative POSIX path")
        raw = unquote(parsed.path)
        base = PurePosixPath(self.source_logical_path).parent
        candidate = base.joinpath(PurePosixPath(raw))
        if any(part in {"", ".", ".."} for part in candidate.parts):
            raise DeclaredResourceError("resource target is not normalized")
        logical_path = candidate.as_posix()
        resolved = self.resources.get(logical_path)
        if resolved is None:
            raise DeclaredResourceError(f"undeclared linked resource: {logical_path}")
        return resolved


def bind_declared_markdown_images(text: str, resolver: DeclaredResourceResolver) -> str:
    """Replace standard Markdown image destinations with declared physical copies."""
    from docwen_core.links._markdown_inline import parse_inline_link, parse_markdown_destination

    def bind(segment: str) -> str:
        parts: list[str] = []
        cursor = 0
        index = 0
        while index < len(segment):
            construct = parse_inline_link(segment, index, image=True)
            if construct is None:
                index += 1
                continue
            parsed = parse_markdown_destination(construct.target, allow_image_size=True)
            if parsed is None:
                raise DeclaredResourceError("invalid Markdown image destination")
            physical_path = resolver.resolve(parsed.destination)
            escaped = physical_path.replace("\\", "/")
            replacement = f"![{construct.label}](<{escaped}>{parsed.suffix})"
            parts.extend((segment[cursor:index], replacement))
            cursor = construct.end
            index = construct.end
        parts.append(segment[cursor:])
        return "".join(parts)

    return _map_visible_markdown(text, bind)


def reject_declared_input_link_lookups(
    text: str,
    *,
    wiki_mode: str = "resolve",
    declared_wiki_link: Callable[[str, str], str | None] | None = None,
) -> None:
    """Reject link forms that would make a declared-input task probe local files."""

    from docwen_core.links._markdown_inline import parse_inline_link

    if wiki_mode in {"keep", "extract_text", "remove"}:
        return

    def reject(segment: str) -> str:
        index = 0
        while index < len(segment):
            construct = parse_inline_link(segment, index, image=True)
            if construct is None:
                construct = parse_inline_link(segment, index, image=False)
            if construct is not None:
                index = construct.end
                continue
            match = _WIKI_LINK_RE.match(segment, index)
            if match is None:
                index += 1
                continue
            index = match.end()
            if match.group(0).startswith("!"):
                continue
            authored_token = match.group(0)
            target = authored_token[2:-2].split("|", 1)[0]
            if target.startswith("#") or urlsplit(target).scheme.lower() in _REMOTE_LINK_SCHEMES:
                continue
            if declared_wiki_link is not None and declared_wiki_link(authored_token, target) is not None:
                continue
            raise DeclaredResourceError("local wiki link resolution is unavailable for declared-input requests")
        return segment

    _map_visible_markdown(text, reject)
