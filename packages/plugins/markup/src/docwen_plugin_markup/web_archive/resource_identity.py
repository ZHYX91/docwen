"""MHTML resource identity is a URI or CID, never an output filename."""

from __future__ import annotations

from urllib.parse import unquote, urlsplit


def _query_defined(uri: str) -> bool:
    return "?" in uri.partition("#")[0]


def _authority_defined(uri: str) -> bool:
    scheme = urlsplit(uri).scheme
    return (uri[len(scheme) + 1 :] if scheme else uri).startswith("//")


def _compose_uri(
    scheme: str,
    authority: str,
    path: str,
    query: str,
    *,
    query_defined: bool,
    authority_defined: bool = False,
    fragment: str = "",
    fragment_defined: bool = False,
) -> str:
    # urlunsplit invents an empty authority for rootless HTTP paths. Keep
    # absent/empty authority and rootless/rooted path separate (RFC 3986 5.3).
    uri = (scheme + ":" if scheme else "") + ("//" + authority if authority or authority_defined else "") + path
    if query_defined:
        uri += "?" + query
    if fragment_defined:
        uri += "#" + fragment
    return uri


def _remove_dot_segments(path: str) -> str:
    """RFC 3986 5.2.4, without folding ordinary empty or escaped segments."""
    output: list[str] = []
    cursor, end = 0, len(path)
    while cursor < end:
        if path.startswith("../", cursor):
            cursor += 3
        elif path.startswith(("./", "/./"), cursor):
            cursor += 2
        elif path.startswith("/../", cursor):
            cursor += 3
            if output:
                output.pop()
        elif (cursor + 2 == end and path.startswith("/.", cursor)) or (
            cursor + 3 == end and path.startswith("/..", cursor)
        ):
            if path.startswith("/..", cursor) and output:
                output.pop()
            output.append("/")
            break
        elif (cursor + 1 == end and path[cursor] == ".") or (cursor + 2 == end and path.startswith("..", cursor)):
            break
        else:
            boundary = path.find("/", cursor + 1 if path[cursor] == "/" else cursor)
            if boundary == -1:
                boundary = end
            output.append(path[cursor:boundary])
            cursor = boundary
    return "".join(output)


def resolve_mhtml_uri(base_uri: str, reference: str) -> str:
    """Merge URI components locally, preserving synthetic and empty-path identity."""
    base_uri = base_uri or "thismessage:/"
    base, ref = urlsplit(base_uri), urlsplit(reference)
    authority_defined = _authority_defined(reference)
    scheme, authority, path = ref.scheme, ref.netloc, ref.path
    query, query_defined = ref.query, _query_defined(reference)
    if not ref.scheme:
        scheme = base.scheme
        if not authority_defined:
            authority, authority_defined = base.netloc, _authority_defined(base_uri)
            if not path:
                path = base.path
                if not query_defined:
                    query, query_defined = base.query, _query_defined(base_uri)
            else:
                if not path.startswith("/"):
                    prefix = "/" if authority_defined and not base.path else base.path.rpartition("/")[0]
                    if base.path and "/" in base.path:
                        prefix += "/"
                    path = prefix + path
                path = _remove_dot_segments(path)
    # Existing explicit scheme/authority identities remain literal. Only a
    # merged relative path has its complete dot segments processed; generic
    # normalization must not collapse a resource's authored path identity.
    return _compose_uri(
        scheme,
        authority,
        path,
        query,
        query_defined=query_defined,
        authority_defined=authority_defined,
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
    return _compose_uri(
        uri.scheme.lower(),
        authority,
        uri.path,
        uri.query,
        query_defined=_query_defined(resolved),
        authority_defined=_authority_defined(resolved),
    )
