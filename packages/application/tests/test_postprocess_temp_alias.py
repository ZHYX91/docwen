"""Owned intermediate files remain admissible beneath a system temp alias."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest
from docx import Document

from docwen_application.postprocess_pipeline import execute_postprocessed_request
from docwen_core.detection import FileAdmissionPathError, enforce_file_admission, inspect_file
from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION, ConversionRequest, OutputPolicy
from docwen_core.models.result import ConversionResult

pytestmark = pytest.mark.unit


@pytest.mark.skipif(os.name == "nt", reason="POSIX system temp alias, including macOS /var")
def test_proofread_uses_physical_owned_temp_but_rejects_linked_user_input(tmp_path, monkeypatch):
    physical = tmp_path / "physical"
    physical.mkdir()
    alias = tmp_path / "alias"
    alias.symlink_to(physical, target_is_directory=True)
    monkeypatch.setattr(tempfile, "tempdir", str(alias))
    source = physical / "note.md"
    source.write_text("# Original\nProofread this document.\n", encoding="utf-8")
    original = source.read_bytes()
    with pytest.raises(FileAdmissionPathError):
        inspect_file(str(alias / source.name))

    stages = []

    def execute(stage, task_id):
        stages.append(enforce_file_admission(stage))
        if not stage.action_name:
            output = Path(stage.output_policy.output_dir) / "rendered.docx"
            document = Document()
            document.add_paragraph("Original document")
            document.save(str(output))
            return ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="rendered",
                        kind="primary",
                        staging_path=str(output),
                        suggested_name=output.name,
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        is_primary=True,
                    )
                ],
            )
        assert stage.action_name == "validate"
        return ConversionResult(task_id=task_id, success=True)

    result = execute_postprocessed_request(
        ConversionRequest(
            request_id="temp-alias",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="docx",
            options={POSTPROCESS_PROOFREAD_OPTION: {"enable_typos_rule": True}},
            output_policy=OutputPolicy(output_dir=str(physical / "published")),
        ),
        "temp-alias",
        execute,
    )
    assert result.success
    assert len(stages) == 2
    assert not Path(stages[1].input_refs[0].path).exists()
    assert not list(physical.glob("docwen_postprocess_*"))
    assert source.read_bytes() == original
