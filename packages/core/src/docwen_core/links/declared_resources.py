"""Fail-closed resolution inside a request-declared virtual input root."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import PurePosixPath
from urllib.parse import unquote, urlsplit

from docwen_core.links._non_embed import _map_visible_markdown

_WIKI_LINK_RE = re.compile(r"!?\[\[(?:[^\]\\]|\\.)+\]\]")
_REMOTE_LINK_SCHEMES = frozenset({"ftp", "http", "https", "mailto"})


class DeclaredResourceError(ValueError):
    """Raised when a local resource is not a declared request input."""


@dataclass(frozen=True, slots=True)
class DeclaredResourceResolver:
    """Resolve logical Markdown targets without consulting the physical filesystem."""

    source_logical_path: str
    resources: dict[str, str]

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


def reject_declared_input_link_lookups(text: str) -> None:
    """Reject link forms that would make a declared-input task probe local files."""

    from docwen_core.links._markdown_inline import parse_inline_link, parse_markdown_destination

    def reject(segment: str) -> str:
        if _WIKI_LINK_RE.search(segment):
            raise DeclaredResourceError("wiki links and embeds are unavailable for declared-input requests")
        index = 0
        while index < len(segment):
            construct = parse_inline_link(segment, index, image=False)
            if construct is None:
                index += 1
                continue
            parsed = parse_markdown_destination(construct.target)
            if parsed is None:
                raise DeclaredResourceError("invalid Markdown link destination")
            destination = unquote(parsed.destination)
            if destination.startswith("#"):
                index = construct.end
                continue
            target = urlsplit(destination)
            if target.scheme.lower() in _REMOTE_LINK_SCHEMES or target.netloc:
                index = construct.end
                continue
            raise DeclaredResourceError("local Markdown links are unavailable for declared-input requests")
        return segment

    _map_visible_markdown(text, reject)
