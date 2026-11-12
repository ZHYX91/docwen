"""Keep the native-acceptance corpus reproducible without simulating Word updates."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from docx import Document
from tests.support.native.docx_seq_acceptance import create_fixture, field_snapshot

from docwen_core.docx_caption_sequences import caption_sequence_identifier
from docwen_core.docx_semantics_v3 import DocxSemanticsV3Recovery

pytestmark = pytest.mark.contract


def test_seq_fixture_separates_generated_ownership_from_native_update_evidence(tmp_path: Path) -> None:
    directory = tmp_path / "native-seq-fixture"
    outputs = create_fixture(directory)
    for name, output in outputs.items():
        observed = field_snapshot(output)["fields"]
        sequences = [field for field in observed if field["instruction"].strip().startswith("SEQ ")]
        references = [field for field in observed if field["instruction"].strip().startswith("REF ")]
        assert [field["cached_result"] for field in sequences] == (
            ["1", "2"] if name == "control" else ["1", "1", "2", "2"]
        )
        assert [field["cached_result"] for field in references] == ["1", "2"]
        recovered = DocxSemanticsV3Recovery.load(output, Document(str(output)))
        assert recovered.paired_caption_map is not None
        counter = caption_sequence_identifier(recovered.paired_caption_map.sequence_scope, "table")
        expected = [f"SEQ {counter} \\* ARABIC"] * 2
        if name != "control":
            expected = ["SEQ Table \\* ARABIC", *expected, "SEQ Table \\* ARABIC"]
        assert [field["instruction"].strip() for field in sequences] == expected
        assert [(item[0], item[1], item[3]) for item in recovered.caption_signatures] == [
            ("table", "generated-first", "1"),
            ("table", "generated-second", "2"),
        ]
    manifest = json.loads((directory / "before-update.json").read_text(encoding="utf-8"))
    assert manifest["native_acceptance"] == "pending"
    with pytest.raises(FileExistsError):
        create_fixture(directory)
