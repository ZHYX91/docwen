"""Published document nodes retain the selected Windows parent's access policy."""

from __future__ import annotations

import os
from datetime import UTC, datetime
from pathlib import Path

import pytest

from docwen_core.models import ArtifactOwnership, OutputPolicy
from docwen_core.models.document_node import ConversionIdentity
from docwen_core.tempfiles import make_temp_dir
from docwen_runtime.output.finalizer import OutputFinalizer

from ._document_node_output_support import _artifact

pytestmark = [
    pytest.mark.integration,
    pytest.mark.skipif(os.name != "nt", reason="Native Windows directory DACL inheritance"),
]


def _security(path: Path):
    import win32security

    descriptor = win32security.GetNamedSecurityInfo(
        str(path),
        win32security.SE_FILE_OBJECT,
        win32security.OWNER_SECURITY_INFORMATION
        | win32security.GROUP_SECURITY_INFORMATION
        | win32security.DACL_SECURITY_INFORMATION,
    )
    acl = descriptor.GetSecurityDescriptorDacl()
    assert acl is not None
    return descriptor, acl


def _descriptor_text(path: Path) -> str:
    import win32security

    return win32security.ConvertSecurityDescriptorToStringSecurityDescriptor(
        _security(path)[0],
        win32security.SDDL_REVISION_1,
        win32security.OWNER_SECURITY_INFORMATION
        | win32security.GROUP_SECURITY_INFORMATION
        | win32security.DACL_SECURITY_INFORMATION,
    )


def _inherited_allow(path: Path, sid: str, required: int) -> bool:
    import win32security

    _, acl = _security(path)
    allowed = 0
    for index in range(acl.GetAceCount()):
        (kind, flags), mask, principal = acl.GetAce(index)
        if (
            kind == win32security.ACCESS_ALLOWED_ACE_TYPE
            and flags & win32security.INHERITED_ACE
            and not flags & win32security.INHERIT_ONLY_ACE
            and win32security.ConvertSidToStringSid(principal) == sid
        ):
            allowed |= mask
    return allowed & required == required


def _shared_parent(tmp_path: Path) -> tuple[Path, str, str, int]:
    import ntsecuritycon
    import win32api
    import win32security

    parent = tmp_path / "selected-parent"
    parent.mkdir()
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32security.TOKEN_QUERY)
    try:
        user = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
    finally:
        win32api.CloseHandle(token)
    users = win32security.CreateWellKnownSid(win32security.WinBuiltinUsersSid)
    _, acl = _security(parent)
    inherit = win32security.CONTAINER_INHERIT_ACE | win32security.OBJECT_INHERIT_ACE
    read_execute = ntsecuritycon.FILE_GENERIC_READ | ntsecuritycon.FILE_GENERIC_EXECUTE
    acl.AddAccessAllowedAceEx(win32security.ACL_REVISION_DS, inherit, ntsecuritycon.FILE_ALL_ACCESS, user)
    acl.AddAccessAllowedAceEx(win32security.ACL_REVISION_DS, inherit, read_execute, users)
    # The only ACL write is preparation of this new, synthetic test-owned parent.
    win32security.SetNamedSecurityInfo(
        str(parent),
        win32security.SE_FILE_OBJECT,
        win32security.DACL_SECURITY_INFORMATION,
        None,
        None,
        acl,
        None,
    )
    return (
        parent,
        win32security.ConvertSidToStringSid(user),
        win32security.ConvertSidToStringSid(users),
        read_execute,
    )


def test_published_node_and_descendants_inherit_the_selected_parent_acl(tmp_path: Path) -> None:
    import ntsecuritycon
    import win32security

    parent, user, users, read_execute = _shared_parent(tmp_path)
    before = _descriptor_text(parent)
    private = Path(make_temp_dir(prefix="private-", dir=parent))
    assert _security(private)[0].GetSecurityDescriptorControl()[0] & win32security.SE_DACL_PROTECTED
    assert not _inherited_allow(private, users, read_execute)

    source = tmp_path / "source.docx"
    source.write_bytes(b"synthetic source identity")
    main = tmp_path / "main.md"
    child = tmp_path / "child.md"
    asset = tmp_path / "asset.png"
    main.write_bytes(b"[Annex](annex.md)\n")
    child.write_bytes(b"![Image](asset.png)\n")
    asset.write_bytes(b"synthetic image bytes")
    inputs = {path: path.read_bytes() for path in (source, main, child, asset)}
    result = OutputFinalizer().finalize(
        "permission.inheritance",
        [
            _artifact(main, artifact_id="main", suggested_name="main.md", media_type="text/markdown", primary=True),
            _artifact(
                child,
                artifact_id="child",
                suggested_name="annex.md",
                media_type="text/markdown",
                ownership=ArtifactOwnership("attachment_of", "main", "attachment", 0, "Annex"),
            ),
            _artifact(
                asset,
                artifact_id="asset",
                suggested_name="asset.png",
                media_type="image/png",
                ownership=ArtifactOwnership("resource_of", "child", "image", 0),
            ),
        ],
        OutputPolicy(output_dir=str(parent), overwrite_mode="error"),
        input_path=str(source),
    )
    assert result.success is True, result.diagnostics
    root = Path(result.metrics.extra["document_node_root"])
    published = [root, *root.rglob("*")]
    assert any(path.is_dir() and path != root for path in published)
    assert any(path.suffix == ".png" for path in published)
    assert any(path.suffix == ".md" and path.parent != root for path in published)
    for path in published:
        assert not _security(path)[0].GetSecurityDescriptorControl()[0] & win32security.SE_DACL_PROTECTED, path
        assert _inherited_allow(path, user, ntsecuritycon.FILE_ALL_ACCESS), path
        assert _inherited_allow(path, users, read_execute), path
    assert _descriptor_text(parent) == before
    assert all(path.read_bytes() == content for path, content in inputs.items())
    published_asset = next(item for item in result.artifacts if item.artifact_id == "asset")
    assert Path(published_asset.staging_path).read_bytes() == inputs[asset]


def test_identical_skip_does_not_change_existing_publication_acl(tmp_path: Path) -> None:
    parent, _, _, _ = _shared_parent(tmp_path)
    source = tmp_path / "source.docx"
    source.write_bytes(b"unchanged synthetic input")
    markdown = tmp_path / "body.md"
    markdown.write_bytes(b"unchanged Markdown\n")
    artifact = _artifact(
        markdown, artifact_id="main", suggested_name="body.md", media_type="text/markdown", primary=True
    )
    identity = ConversionIdentity.create(
        task_id="permission.skip",
        source_stem="source",
        source_format="docx",
        created_at=datetime(2026, 11, 12, tzinfo=UTC),
    )
    finalizer = OutputFinalizer()
    first = finalizer.finalize(
        identity.task_id,
        [artifact],
        OutputPolicy(output_dir=str(parent), overwrite_mode="error"),
        input_path=str(source),
        identity=identity,
    )
    assert first.success, first.diagnostics
    root = Path(first.metrics.extra["document_node_root"])
    paths = [parent, root, *root.rglob("*")]
    original_acls = {path: _descriptor_text(path) for path in paths}
    reused = finalizer.finalize(
        identity.task_id,
        [artifact],
        OutputPolicy(output_dir=str(parent), overwrite_mode="skip"),
        input_path=str(source),
        identity=identity,
    )
    assert reused.success and reused.metrics.output_bytes == 0
    assert reused.artifacts[0].metadata["document_node_reused"] is True
    assert {path: _descriptor_text(path) for path in paths} == original_acls
    assert source.read_bytes() == b"unchanged synthetic input"
    assert markdown.read_bytes() == b"unchanged Markdown\n"
