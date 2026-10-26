"""MHTML resource identity is a URI or CID, never an output filename."""

from __future__ import annotations

from urllib.parse import unquote, urljoin, urlsplit, urlunsplit


def mhtml_resource_identity(reference: str, *, base_uri: str = "") -> str:
    reference = reference.strip()
    if reference.lower().startswith("cid:"):
        return "cid:" + unquote(reference[4:]).strip().strip("<>")
    uri = urlsplit(urljoin(base_uri, reference))
    # URI schemes and DNS names are insensitive to case; paths and queries
    # are not. Fragments identify positions within the same MIME resource.
    credentials, separator, host = uri.netloc.rpartition("@")
    authority = credentials + separator + host.lower() if separator else uri.netloc.lower()
    return urlunsplit((uri.scheme.lower(), authority, uri.path, uri.query, ""))
