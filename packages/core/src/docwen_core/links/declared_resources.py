"""Fail-closed resolution inside a request-declared virtual input root."""

from __future__ import annotations

import hashlib
import re
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
    },
}

_WIKI_LINK_RE = re.compile(r"!?\[\[(?:[^\]\\]|\\.)+\]\]")
_REMOTE_LINK_SCHEMES = frozenset({"ftp", "http", "https", "mailto"})


class DeclaredResourceError(ValueError):
    """Raised when a local resource is not a declared request input."""


@dataclass(frozen=True, slots=True)
class DeclaredResourceResolver:
    """Resolve logical Markdown targets without consulting the physical filesystem."""

    source_logical_path: str
    resources: dict[str, str]
    bindings: dict[str, str] = field(default_factory=dict)

    def with_bindings(self, source: str, value: object) -> DeclaredResourceResolver:
        """Authenticate token bindings against this source and the declared inventory."""
        if value is None:
            return self
        if not isinstance(value, dict) or set(value) != {"authored_sha256", "images"}:
            raise DeclaredResourceError("invalid Markdown resource bindings")
        if value["authored_sha256"] != hashlib.sha256(source.encode("utf-8")).hexdigest():
            raise DeclaredResourceError("Markdown resource bindings source hash mismatch")
        images = value["images"]
        if not isinstance(images, list):
            raise DeclaredResourceError("invalid Markdown image bindings")
        visible_tokens: set[str] = set()

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
        return DeclaredResourceResolver(self.source_logical_path, self.resources, bindings)

    def resolve_image(self, authored_token: str, target: str) -> str:
        logical = self.bindings.get(authored_token)
        return self.resources[logical] if logical is not None else self.resolve(target)

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


def reject_declared_input_link_lookups(text: str, *, wiki_mode: str = "resolve") -> None:
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
            target = match.group(0)[2:-2].split("|", 1)[0]
            if target.startswith("#") or urlsplit(target).scheme.lower() in _REMOTE_LINK_SCHEMES:
                continue
            raise DeclaredResourceError("local wiki link resolution is unavailable for declared-input requests")
        return segment

    _map_visible_markdown(text, reject)
