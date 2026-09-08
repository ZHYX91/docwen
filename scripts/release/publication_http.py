"""Bounded read recovery; write outcomes are reconciled without repeating writes."""

from __future__ import annotations

import http.client
import json
import os
import re
import subprocess
import time
import urllib.error
import urllib.request
from collections.abc import Callable
from contextlib import suppress
from datetime import UTC, datetime
from email.utils import parsedate_to_datetime
from typing import Any

from scripts.release.publication_contract import PublicationError


class PendingRead(PublicationError):
    def __init__(self, message: str, *, retry_after: float = 0) -> None:
        super().__init__(message)
        self.retry_after = retry_after


class ApiError(PublicationError):
    def __init__(self, status: int, *, retry_after: float = 0, limited: bool = False) -> None:
        super().__init__(f"GitHub request failed: HTTP {status}")
        self.status = status
        self.retry_after = retry_after
        self.transient = status in {408, 429, 500, 502, 503, 504} or (status == 403 and limited)


def retry_after_seconds(value: str | None) -> float:
    if not value:
        return 0
    try:
        return max(0, float(value))
    except ValueError:
        try:
            return max(0, (parsedate_to_datetime(value) - datetime.now(UTC)).total_seconds())
        except (TypeError, ValueError):
            return 0


def read_with_retry[T](
    operation: Callable[[float], T],
    *,
    budget: float = 60,
    clock: Callable[[], float] = time.monotonic,
    sleep: Callable[[float], None] = time.sleep,
    allow_missing: bool = False,
) -> T:
    deadline = clock() + budget
    attempt = 0
    while True:
        remaining = deadline - clock()
        if remaining <= 0:
            raise PublicationError("read recovery budget exhausted")
        try:
            result = operation(remaining)
            if clock() > deadline:
                raise PublicationError("read recovery budget exhausted during request")
            return result
        except (PendingRead, ApiError, TimeoutError, OSError, http.client.HTTPException, json.JSONDecodeError) as error:
            if isinstance(error, ApiError) and not (error.transient or (allow_missing and error.status == 404)):
                raise
            remaining = deadline - clock()
            delay = max(min(2 ** (attempt + 1), 16), getattr(error, "retry_after", 0))
            if delay >= remaining:
                raise PublicationError("read recovery budget exhausted") from error
            sleep(delay)
            attempt += 1


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        # JSON endpoints never require redirecting the bearer token.
        raise PublicationError("unexpected authenticated API redirect")


class GitHub:
    def __init__(self, repository: str, *, token: str | None = None) -> None:
        self.repository = repository
        token = token or os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
        if not token:
            result = subprocess.run(["gh", "auth", "token"], check=True, capture_output=True, text=True, timeout=30)
            token = result.stdout.strip()
        if not token:
            raise PublicationError("GitHub authentication unavailable")
        self._token = token
        self._opener = urllib.request.build_opener(_NoRedirect())
        self._not_before = 0.0

    def request(self, method: str, path: str, *, timeout: float, body: object = None, data: bytes | None = None) -> Any:
        remaining_limit = self._not_before - time.monotonic()
        if remaining_limit > 0:
            raise PendingRead("waiting for GitHub rate limit", retry_after=remaining_limit)
        host = "https://uploads.github.com" if data is not None else "https://api.github.com"
        if not path.startswith(f"/repos/{self.repository}/"):
            raise PublicationError("request outside the publication repository")
        headers = {
            "Authorization": f"Bearer {self._token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2026-03-10",
            "User-Agent": "DocWen-release",
        }
        if body is not None:
            data = json.dumps(body).encode("utf-8")
            headers["Content-Type"] = "application/json"
        elif data is not None:
            headers["Content-Type"] = "application/octet-stream"
        request = urllib.request.Request(host + path, data=data, headers=headers, method=method)
        try:
            with self._opener.open(request, timeout=timeout) as response:
                return json.load(response)
        except urllib.error.HTTPError as error:
            retry_after = retry_after_seconds(error.headers.get("Retry-After"))
            limited = bool(retry_after) or error.headers.get("X-RateLimit-Remaining") == "0"
            if limited and not retry_after:
                with suppress(ValueError):
                    retry_after = max(0, float(error.headers.get("X-RateLimit-Reset", "0")) - time.time())
            if limited:
                self._not_before = time.monotonic() + retry_after
            raise ApiError(error.code, retry_after=retry_after, limited=limited) from None

    def get(self, path: str, *, allow_missing: bool = False) -> Any:
        return read_with_retry(lambda timeout: self.request("GET", path, timeout=timeout), allow_missing=allow_missing)

    def download_identity(self, path: str, *, timeout: float) -> dict[str, Any]:
        import hashlib

        try:
            result = subprocess.run(
                ["gh", "api", path, "-H", "Accept: application/octet-stream"],
                capture_output=True,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as error:
            raise PendingRead("asset download timed out") from error
        if result.returncode:
            match = re.search(rb"HTTP (\d{3})", result.stderr)
            if match:
                raise ApiError(int(match[1]))
            raise PendingRead("asset download failed")
        return {"bytes": len(result.stdout), "sha256": hashlib.sha256(result.stdout).hexdigest()}
