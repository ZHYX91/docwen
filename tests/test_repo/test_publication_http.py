from __future__ import annotations

import io
import urllib.request

import pytest
from scripts.release.publication_contract import PublicationError
from scripts.release.publication_http import GitHub

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("path", ["/repos/owner/product", "/repos/owner/product/commits/main"])
def test_repository_metadata_and_child_requests_reach_transport(monkeypatch: pytest.MonkeyPatch, path: str) -> None:
    monkeypatch.setenv("GH_TOKEN", "test-only-placeholder")
    api = GitHub("owner/product")
    requests: list[tuple[str, float]] = []

    def open_request(request: urllib.request.Request, *, timeout: float) -> io.BytesIO:
        requests.append((request.full_url, timeout))
        return io.BytesIO(b'{"default_branch":"main"}')

    monkeypatch.setattr(api._opener, "open", open_request)
    assert api.request("GET", path, timeout=3) == {"default_branch": "main"}
    assert requests == [(f"https://api.github.com{path}", 3)]


@pytest.mark.parametrize(
    "path",
    ["/repos/owner/product-extra", "/repos/owner/product-extra/releases", "/repos/other/product", "/user"],
)
def test_repository_boundary_rejects_other_resources_before_transport(
    monkeypatch: pytest.MonkeyPatch, path: str
) -> None:
    monkeypatch.setenv("GH_TOKEN", "test-only-placeholder")
    api = GitHub("owner/product")

    def unexpected_request(*args: object, **kwargs: object) -> None:
        pytest.fail("out-of-scope request reached transport")

    monkeypatch.setattr(api._opener, "open", unexpected_request)
    with pytest.raises(PublicationError, match="outside the publication repository"):
        api.request("GET", path, timeout=3)
