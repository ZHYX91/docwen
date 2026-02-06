from __future__ import annotations

import http.client
import socket
import ssl
import urllib.request
import ftplib
import smtplib

import pytest

import docwen.security.network_isolation as ni


@pytest.mark.unit
def test_blocked_socket_raises() -> None:
    with pytest.raises(ni.NetworkIsolationError):
        ni.BlockedSocket()


@pytest.mark.unit
def test_get_network_isolation_status_reflects_patched_modules() -> None:
    orig_socket = socket.socket
    orig_http = http.client.HTTPConnection
    orig_urlopen = urllib.request.urlopen
    orig_ftp = ftplib.FTP
    orig_smtp = smtplib.SMTP
    orig_wrap_socket = getattr(ssl, "wrap_socket", None)

    try:
        socket.socket = ni.BlockedSocket
        http.client.HTTPConnection = ni.BlockedHTTPConnection

        class _BlockedUrlopen:
            def __call__(self, *args, **kwargs):
                raise ni.NetworkIsolationError("x")

            def __str__(self) -> str:
                return "throw"

        urllib.request.urlopen = _BlockedUrlopen()

        class _BlockedFTP:
            pass

        class _BlockedSMTP:
            pass

        ftplib.FTP = _BlockedFTP
        smtplib.SMTP = _BlockedSMTP

        def _wrap_socket(*args, **kwargs):
            raise ni.NetworkIsolationError("x")

        ssl.wrap_socket = _wrap_socket

        status = ni.get_network_isolation_status()
        assert status["socket_blocked"] is True
        assert status["http_blocked"] is True
        assert status["urllib_blocked"] is True
        assert status["ftp_blocked"] is True
        assert status["smtp_blocked"] is True
        assert status["ssl_blocked"] is True
    finally:
        socket.socket = orig_socket
        http.client.HTTPConnection = orig_http
        urllib.request.urlopen = orig_urlopen
        ftplib.FTP = orig_ftp
        smtplib.SMTP = orig_smtp
        if orig_wrap_socket is not None:
            ssl.wrap_socket = orig_wrap_socket
