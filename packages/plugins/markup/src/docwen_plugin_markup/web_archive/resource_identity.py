"""MHTML resource identity is a URI or CID, never an output filename."""

from __future__ import annotations

from urllib.parse import unquote, urljoin, urlsplit, urlunsplit


def _query_defined(uri: str) -> bool:
    return "?" in uri.partition("#")[0]


def _compose_uri(
    scheme: str,
    authority: str,
    path: str,
    query: str,
    *,
    query_defined: bool,
    fragment: str = "",
    fragment_defined: bool = False,
) -> str:
    uri = urlunsplit((scheme, authority, path, query, ""))
    if query_defined and not query:
        uri += "?"
    if fragment_defined:
        uri += "#" + fragment
    return uri


def resolve_mhtml_uri(base_uri: str, reference: str) -> str:
    """Resolve even RFC 2557's synthetic base, which urljoin does not know."""
    base_uri = base_uri or "thismessage:/"
    base = urlsplit(base_uri)
    ref = urlsplit(reference)
    synthetic = base.scheme == "thismessage" and not ref.scheme
    join_base = base_uri
    if synthetic:
        join_base = _compose_uri(
            "https",
            base.netloc or "docwen-mhtml.invalid",
            base.path or "/",
            base.query,
            query_defined=_query_defined(base_uri),
        )
    resolved = urlsplit(urljoin(join_base, reference))
    query_defined = _query_defined(reference)
    query = ref.query
    if not (ref.scheme or ref.netloc or ref.path or query_defined):
        query, query_defined = base.query, _query_defined(base_uri)
    if synthetic:
        authority = resolved.netloc if base.netloc or ref.netloc else ""
        scheme = "thismessage"
    else:
        authority, scheme = resolved.netloc, resolved.scheme
    # RFC 3986 5.2.2 distinguishes an absent query from an explicit empty one;
    # urllib's urljoin alone treats '?' as though its query were absent.
    return _compose_uri(
        scheme,
        authority,
        resolved.path,
        query,
        query_defined=query_defined,
        fragment=ref.fragment,
        fragment_defined="#" in reference,
    )


def content_id_identity(header: str) -> str:
    """A Content-ID header is already the literal identifier, never URL-decoded."""
    return "cid:" + header.strip().strip("<>")


def mhtml_resource_identity(reference: str, *, base_uri: str = "") -> str:
    reference = reference.strip()
    if reference.lower().startswith("cid:"):
        return "cid:" + unquote(reference[4:], errors="strict")
    resolved = resolve_mhtml_uri(base_uri, reference)
    uri = urlsplit(resolved)
    # URI schemes and DNS names are insensitive to case; paths and queries
    # are not. Fragments identify positions within the same MIME resource.
    credentials, separator, host = uri.netloc.rpartition("@")
    authority = credentials + separator + host.lower() if separator else uri.netloc.lower()
    return _compose_uri(uri.scheme.lower(), authority, uri.path, uri.query, query_defined=_query_defined(resolved))
