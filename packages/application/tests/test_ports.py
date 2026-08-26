"""Contract tests for application ports (protocols).

Verifies:
- Ports are Protocols (structural subtyping works).
- Port method signatures are correct.
- Ports can be implemented by plain classes (no runtime base class needed).
- Ports are re-exported from docwen_application.ports.
"""

from typing import Protocol

import pytest

pytestmark = pytest.mark.unit


# ── Protocol identity ───────────────────────────────────────────────────────


def test_runtime_port_is_protocol() -> None:
    from docwen_application.ports.runtime import RuntimePort

    assert issubclass(RuntimePort, Protocol)


def test_config_port_is_protocol() -> None:
    from docwen_application.ports.runtime import ConfigPort

    assert issubclass(ConfigPort, Protocol)


def test_config_port_documents_transactional_batch_and_reset_contract() -> None:
    from docwen_application.ports.runtime import ConfigPort

    for method_name in (
        "set",
        "set_many",
        "save_file_text",
        "reset_file",
        "reset_group",
        "reset_all",
    ):
        method_doc = getattr(ConfigPort, method_name).__doc__ or ""
        assert "all-or-nothing" in method_doc
        assert "compensation succeeds" in method_doc
    assert "untrusted" in (ConfigPort.get.__doc__ or "")
    assert "last reconciliation did not complete" in (ConfigPort.snapshot.__doc__ or "")
    assert "get`` and ``snapshot`` may raise" in (ConfigPort.set_many.__doc__ or "")


def test_presenter_port_is_protocol() -> None:
    from docwen_application.ports.runtime import PresenterPort

    assert issubclass(PresenterPort, Protocol)


def test_output_manifest_persistence_port_is_optional_protocol() -> None:
    from docwen_application.ports.runtime import OutputManifestPersistencePort

    class PlainPersistence:
        def persist_output_manifests(self, request, result):
            del request
            return result

    assert issubclass(OutputManifestPersistencePort, Protocol)
    assert isinstance(PlainPersistence(), OutputManifestPersistencePort)


# ── Ports package re-exports ────────────────────────────────────────────────


def test_ports_init_exports_all_ports() -> None:
    from docwen_application.ports import ConfigPort, OutputManifestPersistencePort, PresenterPort, RuntimePort

    assert RuntimePort is not None
    assert ConfigPort is not None
    assert PresenterPort is not None
    assert OutputManifestPersistencePort is not None


# ── Structural subtyping: classes WITHOUT explicit Protocol base ────────────


def test_plain_class_satisfies_runtime_port():
    """A plain class with the required methods satisfies RuntimePort.

    Note: issubclass() is not used here because RuntimePort has a @property
    (is_available) which makes it a "protocol with non-method members".
    Python 3.12's @runtime_checkable does not support issubclass() for such
    protocols. isinstance() works correctly (it checks only callable members).
    """
    from docwen_application.ports.runtime import RuntimePort

    class FakeRuntime:
        def execute(self, request):
            return object()

        def cancel(self, task_id: str) -> None:
            pass

        def shutdown(self) -> None:
            pass

        @property
        def is_available(self) -> bool:
            return True

    assert isinstance(FakeRuntime(), RuntimePort)


def test_plain_class_satisfies_config_port():
    """A plain class with required methods satisfies ConfigPort."""
    from docwen_application.ports.runtime import ConfigPort

    class FakeConfig:
        def get(self, key: str, default=None):
            return default

        def snapshot(self) -> dict[str, object]:
            return {}

        def set(self, key: str, value: object) -> bool:
            return True

        def set_many(self, values: dict[str, object]) -> bool:
            return True

        def get_file_text(self, rel_path: str) -> str | None:
            return ""

        def save_file_text(self, rel_path: str, content: str) -> bool:
            return True

        def reset_file(self, rel_path: str) -> bool:
            return True

        def reset_group(self, group: str) -> bool:
            return True

        def reset_all(self) -> bool:
            return True

        def reload(self) -> None:
            return None

    assert issubclass(FakeConfig, ConfigPort)
    assert isinstance(FakeConfig(), ConfigPort)


def test_plain_class_satisfies_presenter_port():
    """A plain class with present_result/present_error satisfies PresenterPort."""
    from docwen_application.ports.runtime import PresenterPort

    class FakePresenter:
        def present_result(self, result) -> None:
            pass

        def present_error(self, task_id: str, error: Exception) -> None:
            pass

    assert issubclass(FakePresenter, PresenterPort)
    assert isinstance(FakePresenter(), PresenterPort)


# ── Protocol method validation ──────────────────────────────────────────────


def test_runtime_port_required_methods() -> None:
    """Verify RuntimePort exposes execute/cancel/shutdown/is_available.

    Uses isinstance() rather than issubclass() because RuntimePort has a
    @property member, which Python 3.12 @runtime_checkable does not support
    with issubclass().
    """
    from docwen_application.ports.runtime import RuntimePort

    # A class missing 'cancel' should NOT satisfy the protocol
    class MissingCancel:
        def execute(self, request):
            pass

        def shutdown(self) -> None:
            pass

        @property
        def is_available(self) -> bool:
            return True

    assert not isinstance(MissingCancel(), RuntimePort)

    # A class missing 'is_available' should NOT satisfy the protocol
    class MissingIsAvailable:
        def execute(self, request):
            pass

        def cancel(self, task_id: str) -> None:
            pass

        def shutdown(self) -> None:
            pass

    assert not isinstance(MissingIsAvailable(), RuntimePort)

    class MissingShutdown:
        def execute(self, request):
            pass

        def cancel(self, task_id: str) -> None:
            pass

        @property
        def is_available(self) -> bool:
            return True

    assert not isinstance(MissingShutdown(), RuntimePort)


def test_config_port_required_methods() -> None:
    """Verify ConfigPort requires the configured method surface."""
    from docwen_application.ports.runtime import ConfigPort

    class NoGet:
        pass

    assert not issubclass(NoGet, ConfigPort)

    class OnlyGet:
        def get(self, key: str, default=None):
            return default

    assert not issubclass(OnlyGet, ConfigPort)

    class MissingResetGroup:
        def get(self, key: str, default=None):
            return default

        def snapshot(self) -> dict[str, object]:
            return {}

        def set(self, key: str, value: object) -> bool:
            return True

        def set_many(self, values: dict[str, object]) -> bool:
            return True

        def get_file_text(self, rel_path: str) -> str | None:
            return ""

        def save_file_text(self, rel_path: str, content: str) -> bool:
            return True

        def reset_file(self, rel_path: str) -> bool:
            return True

        def reset_all(self) -> bool:
            return True

        def reload(self) -> None:
            return None

    assert not issubclass(MissingResetGroup, ConfigPort)


def test_presenter_port_required_methods() -> None:
    """Verify PresenterPort exposes present_result and present_error."""
    from docwen_application.ports.runtime import PresenterPort

    class MissingBoth:
        pass

    assert not issubclass(MissingBoth, PresenterPort)

    class MissingError:
        def present_result(self, result) -> None:
            pass

    assert not issubclass(MissingError, PresenterPort)


# ── Mock compatibility ─────────────────────────────────────────────────────


def test_unittest_mock_satisfies_runtime_port_via_spec() -> None:
    """MagicMock(spec=RuntimePort) passes isinstance check (mock is special)."""
    from unittest.mock import MagicMock

    from docwen_application.ports.runtime import RuntimePort

    mock = MagicMock(spec=RuntimePort)
    # MagicMock unconditionally passes isinstance checks — this test documents
    # that the spec= pattern is valid and won't break.
    assert isinstance(mock, RuntimePort)


def test_unittest_mock_satisfies_config_port_via_spec() -> None:
    from unittest.mock import MagicMock

    from docwen_application.ports.runtime import ConfigPort

    mock = MagicMock(spec=ConfigPort)
    assert isinstance(mock, ConfigPort)


def test_unittest_mock_satisfies_presenter_port_via_spec() -> None:
    from unittest.mock import MagicMock

    from docwen_application.ports.runtime import PresenterPort

    mock = MagicMock(spec=PresenterPort)
    assert isinstance(mock, PresenterPort)


# ── Port closure: the same object cannot be two ports at once? ──────────────
# Actually it CAN — structural subtyping allows it.  Verifies that a single
# adapter class can satisfy multiple ports (common in integration adapters).


def test_multi_port_adapter() -> None:
    from docwen_application.ports.runtime import ConfigPort, RuntimePort

    class RuntimeAdapter:
        def execute(self, request):
            return object()

        def cancel(self, task_id: str) -> None:
            pass

        def shutdown(self) -> None:
            pass

        @property
        def is_available(self) -> bool:
            return True

        def get(self, key: str, default=None):
            return default

        def snapshot(self) -> dict[str, object]:
            return {}

        def set(self, key: str, value: object) -> bool:
            return True

        def set_many(self, values: dict[str, object]) -> bool:
            return True

        def get_file_text(self, rel_path: str) -> str | None:
            return ""

        def save_file_text(self, rel_path: str, content: str) -> bool:
            return True

        def reset_file(self, rel_path: str) -> bool:
            return True

        def reset_group(self, group: str) -> bool:
            return True

        def reset_all(self) -> bool:
            return True

        def reload(self) -> None:
            return None

    # Verifies structural typing: one class can satisfy both protocols.
    adapter = RuntimeAdapter()
    assert isinstance(adapter, RuntimePort)
    assert isinstance(adapter, ConfigPort)
