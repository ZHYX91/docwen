from __future__ import annotations

import datetime as dt

import pytest

import docwen.security.expiration_check as expiration_check


@pytest.mark.unit
def test_deobfuscate_invalid_data_returns_fallback() -> None:
    assert expiration_check._deobfuscate("not_base64", "k") == "2000-01-01"


@pytest.mark.unit
def test_get_expiration_status_scenarios(monkeypatch: pytest.MonkeyPatch) -> None:
    now = dt.datetime.now()

    monkeypatch.setattr(expiration_check, "EXPIRATION_DATE", now + dt.timedelta(days=100), raising=True)
    info = expiration_check.get_expiration_status()
    assert info.status == expiration_check.ExpirationStatus.VALID

    monkeypatch.setattr(expiration_check, "EXPIRATION_DATE", now + dt.timedelta(days=15), raising=True)
    info = expiration_check.get_expiration_status()
    assert info.status == expiration_check.ExpirationStatus.NEARING_EXPIRATION

    monkeypatch.setattr(expiration_check, "EXPIRATION_DATE", now - dt.timedelta(days=1), raising=True)
    info = expiration_check.get_expiration_status()
    assert info.status == expiration_check.ExpirationStatus.EXPIRED

