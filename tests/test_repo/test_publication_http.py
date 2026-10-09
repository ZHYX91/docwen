from __future__ import annotations

import io
import urllib.error
import urllib.request
from email.message import Message

import pytest
from scripts.release.publication_contract import PublicationError
from scripts.release.publication_http import GitHub

pytestmark = pytest.mark.unit


def test_artifact_redirect_is_returned_without_sending_credentials_to_storage(monkeypatch):
    monkeypatch.setenv("GH_TOKEN", "test-only-placeholder")
    api = GitHub("owner/product")
    requests = []

    class Opener:
        def open(self, request, *, timeout):
            requests.append(request)
            headers = Message()
            headers["Location"] = "https://storage.invalid/signed"
            raise urllib.error.HTTPError(request.full_url, 302, "Found", headers, io.BytesIO())

    monkeypatch.setattr(urllib.request, "build_opener", lambda handler: Opener())
    assert api.artifact_url(42, timeout=3) == "https://storage.invalid/signed"
    assert len(requests) == 1
    assert requests[0].full_url == "https://api.github.com/repos/owner/product/actions/artifacts/42/zip"
    assert requests[0].get_header("Authorization") == "Bearer test-only-placeholder"


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
