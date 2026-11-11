"""The v4 candidate creator owns failures after the package has been saved."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_core.docx_resolved_numbering import ResolvedNumberingDocxError
from docwen_plugin_markdown.to_docx import resolved_v4_route as route
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_docx.managed_styles import ManagedStyleCompletionError

from ._resolved_v4_support import _NEUTRAL, _PLAN, _context, _refs

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("failure", ["style", "proof", "publish"])
def test_saved_candidate_is_removed_when_late_proof_or_publication_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, failure: str
) -> None:
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))
    untouched = Path(workspace.staging_dir) / "preexisting.txt"
    untouched.write_text("Unrelated content", encoding="utf-8")
    observed: list[int] = []

    def fail_style(*_args):
        candidate = next(Path(workspace.staging_dir).glob("*.candidate"))
        observed.append(candidate.stat().st_size)
        raise ManagedStyleCompletionError(
            "MD2DOCX-STYLE-COMPLETION-ERROR", "injected style proof failure", error_type="conversion_failed"
        )

    original_proof = route.ResolvedNumberingDocxSession.prove_package

    def fail_proof(session, path):
        if Path(path).suffix == ".candidate":
            observed.append(Path(path).stat().st_size)
            raise ResolvedNumberingDocxError("injected final proof failure")
        return original_proof(session, path)

    def fail_publication(source, _destination):
        observed.append(Path(source).stat().st_size)
        raise OSError("injected publication failure")

    if failure == "style":
        monkeypatch.setattr(route, "validate_managed_style_package", fail_style)
    elif failure == "proof":
        monkeypatch.setattr(route.ResolvedNumberingDocxSession, "prove_package", fail_proof)
    else:
        monkeypatch.setattr(route, "replace_authorized_file", fail_publication)
    result = MdToDocxConverter().convert(context)
    assert not result.success and result.error is not None
    assert "injected" in result.error.message
    assert observed and min(observed) > 1000
    assert not workspace.registered_artifacts
    assert list(Path(workspace.staging_dir).iterdir()) == [untouched]
    assert untouched.read_text(encoding="utf-8") == "Unrelated content"


def test_candidate_cleanup_error_does_not_hide_original_proof_diagnostic(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    def fail_style(*_args):
        raise ManagedStyleCompletionError(
            "MD2DOCX-STYLE-COMPLETION-ERROR", "original proof failure", error_type="conversion_failed"
        )

    original_unlink = Path.unlink

    def fail_cleanup(path, missing_ok=False):
        if path.suffix == ".candidate" and missing_ok and path.exists():
            raise PermissionError("injected candidate cleanup failure")
        return original_unlink(path, missing_ok=missing_ok)

    monkeypatch.setattr(route, "validate_managed_style_package", fail_style)
    monkeypatch.setattr(Path, "unlink", fail_cleanup)
    result = MdToDocxConverter().convert(context)
    assert not result.success and result.error is not None
    assert result.error.diagnostic_code == "MD2DOCX-STYLE-COMPLETION-ERROR"
    assert "original proof failure" in result.error.message and "candidate cleanup failed" in result.error.message
    assert any(item.code == "MD2DOCX-RESOLVED-V4-CLEANUP-FAILED" for item in result.diagnostics)
    assert not workspace.registered_artifacts
    candidate = next(Path(workspace.staging_dir).glob("*.candidate"))
    original_unlink(candidate)


def test_preparation_failure_does_not_leave_an_unowned_candidate(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import docwen_core.docx_table_roles as table_roles

    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))
    original_preparation = table_roles.prepare_table_roles

    def fail_preparation(document, **kwargs):
        if kwargs:
            return original_preparation(document, **kwargs)
        raise ValueError("injected table preparation failure")

    monkeypatch.setattr(table_roles, "prepare_table_roles", fail_preparation)
    result = MdToDocxConverter().convert(context)
    assert not result.success and result.error is not None
    assert "injected table preparation failure" in result.error.message
    assert not workspace.registered_artifacts
    assert not list(Path(workspace.staging_dir).iterdir())
