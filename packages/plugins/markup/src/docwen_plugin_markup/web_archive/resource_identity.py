"""MHTML resource identity is a URI or CID, never an output filename."""

from __future__ import annotations

from urllib.parse import unquote, urljoin, urlsplit, urlunsplit


def resolve_mhtml_uri(base_uri: str, reference: str) -> str:
    """Resolve even RFC 2557's synthetic base, which urljoin does not know."""
    base_uri = base_uri or "thismessage:/"
    if base_uri.startswith("thismessage:") and not urlsplit(reference).scheme:
        base = urlsplit(base_uri)
        resolved = urlsplit(urljoin("https://docwen-mhtml.invalid" + (base.path or "/"), reference))
        authority = "" if resolved.netloc == "docwen-mhtml.invalid" else resolved.netloc
        return urlunsplit(("thismessage", authority, resolved.path, resolved.query, resolved.fragment))
    return urljoin(base_uri, reference)


def content_id_identity(header: str) -> str:
    """A Content-ID header is already the literal identifier, never URL-decoded."""
    return "cid:" + header.strip().strip("<>")


def mhtml_resource_identity(reference: str, *, base_uri: str = "") -> str:
    reference = reference.strip()
    if reference.lower().startswith("cid:"):
        return "cid:" + unquote(reference[4:], errors="strict")
    uri = urlsplit(resolve_mhtml_uri(base_uri, reference))
    # URI schemes and DNS names are insensitive to case; paths and queries
    # are not. Fragments identify positions within the same MIME resource.
    credentials, separator, host = uri.netloc.rpartition("@")
    authority = credentials + separator + host.lower() if separator else uri.netloc.lower()
    return urlunsplit((uri.scheme.lower(), authority, uri.path, uri.query, ""))
