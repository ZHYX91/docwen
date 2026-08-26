"""Hash, schema, and corpus gates for the active Markdown v3 source oracle."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from jsonschema import Draft202012Validator
from referencing import Registry, Resource

from docwen_plugin_markdown.document_semantics_v3 import (
    DIAGNOSTICS_SCHEMA_ID,
    ExternalCitationResolution,
    ExternalReferenceResolution,
    analyze_markdown_semantics_v3,
)

pytestmark = pytest.mark.contract

REPO_ROOT = Path(__file__).resolve().parents[2]
ORACLE_ROOT = REPO_ROOT / "contracts" / "oracles" / "docwen.markdown_semantics.v3"
PROJECTION_SCHEMA_PATH = ORACLE_ROOT / "schemas" / "projection.schema.json"
DIAGNOSTICS_SCHEMA_PATH = ORACLE_ROOT / "schemas" / "diagnostics.schema.json"
CASE_SCHEMA_PATH = ORACLE_ROOT / "schemas" / "corpus-case.schema.json"
MANIFEST_SCHEMA_PATH = ORACLE_ROOT / "schemas" / "manifest.schema.json"
MANIFEST_PATH = ORACLE_ROOT / "manifest.json"


def _json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _registry() -> Registry:
    projection = _json(PROJECTION_SCHEMA_PATH)
    return Registry().with_resource(projection["$id"], Resource.from_contents(projection))


def _validate(schema_path: Path, instance: object) -> None:
    schema = _json(schema_path)
    Draft202012Validator.check_schema(schema)
    Draft202012Validator(schema, registry=_registry()).validate(instance)


def _external_reference(record: dict) -> ExternalReferenceResolution:
    copy = dict(record)
    if "heading_path" in copy:
        copy["heading_path"] = tuple(copy["heading_path"])
    return ExternalReferenceResolution(**copy)


def _observed_diagnostic(diagnostic: dict) -> dict:
    observed = {
        "severity": diagnostic["severity"],
        "code": diagnostic["code"],
        "range": diagnostic["range"],
    }
    if "related_ranges" in diagnostic:
        observed["related_ranges"] = diagnostic["related_ranges"]
    if "fixes" in diagnostic:
        observed["fix_ids"] = [item["fix_id"] for item in diagnostic["fixes"]]
    return observed


def test_all_v3_schemas_are_draft_2020_12_and_closed() -> None:
    for path in sorted((ORACLE_ROOT / "schemas").glob("*.schema.json")):
        schema = _json(path)
        Draft202012Validator.check_schema(schema)
        assert schema["$schema"] == "https://json-schema.org/draft/2020-12/schema"
        assert schema["additionalProperties"] is False


def test_every_hash_addressed_corpus_case_matches_runtime_projection_and_diagnostics() -> None:
    case_paths = sorted((ORACLE_ROOT / "corpus").glob("*.case.json"))
    assert len(case_paths) >= 13
    for case_path in case_paths:
        case = _json(case_path)
        _validate(CASE_SCHEMA_PATH, case)
        source = case["source"]
        before = source
        analysis = analyze_markdown_semantics_v3(
            source,
            input_id=case["input_id"],
            external_references=tuple(_external_reference(item) for item in case.get("external_references", [])),
            external_citations=tuple(ExternalCitationResolution(**item) for item in case.get("external_citations", [])),
        )
        _validate(PROJECTION_SCHEMA_PATH, analysis.projection)
        diagnostic_oracle = {
            "$schema": DIAGNOSTICS_SCHEMA_ID,
            "schema": "docwen.markdown_diagnostics.v3",
            "source": analysis.projection["source"],
            "diagnostics": list(analysis.diagnostics),
        }
        _validate(DIAGNOSTICS_SCHEMA_PATH, diagnostic_oracle)

        expected = case["expected"]
        assert [(item["kind"], item.get("id")) for item in analysis.projection["targets"]] == [
            tuple(item) for item in expected["target_kind_ids"]
        ], case["case_id"]
        assert [(item["id"], item["block_kind"], item["placement"]) for item in analysis.projection["anchors"]] == [
            tuple(item) for item in expected["anchor_kind_ids"]
        ], case["case_id"]
        if "anchor_container_paths" in expected:
            assert [item["container_path"] for item in analysis.projection["anchors"]] == expected[
                "anchor_container_paths"
            ], case["case_id"]
        if "fenced_sources" in expected:
            assert analysis.projection["fenced_sources"] == expected["fenced_sources"], case["case_id"]
        assert [item["kind"] for item in analysis.projection["links"]] == expected["link_kinds"], case["case_id"]
        assert [item["resolution_status"] for item in analysis.projection["references"]] == expected[
            "reference_statuses"
        ], case["case_id"]
        assert [
            item["key"] for occurrence in analysis.projection["citations"] for item in occurrence["items"]
        ] == expected["citation_keys"], case["case_id"]
        assert [_observed_diagnostic(item) for item in analysis.diagnostics] == expected["diagnostics"], case["case_id"]
        assert source == before


def test_manifest_binds_final_spec_schema_corpus_and_implementation_bytes() -> None:
    manifest = _json(MANIFEST_PATH)
    _validate(MANIFEST_SCHEMA_PATH, manifest)
    assert manifest["schema"] == "docwen.markdown_semantics_manifest.v1"
    assert manifest["oracle"] == {
        "semantics": "docwen.markdown_semantics.v3",
        "semantics_schema_id": "urn:docwen:schema:markdown-semantics:v3",
        "diagnostics": "docwen.markdown_diagnostics.v3",
        "diagnostics_schema_id": "urn:docwen:schema:markdown-diagnostics:v3",
    }
    assert manifest["final_spec_baseline"] == {
        "commit": "cbb33dba6509c43912bef2df744ad8d8654f628c",
        "tree": "151e15aa1757c9d9a45a2eaab25e6cb101117db8",
    }
    records = manifest["files"]
    assert len(records) == len({item["path"] for item in records})
    assert {item["path"] for item in records} == {
        path.relative_to(ORACLE_ROOT).as_posix()
        for folder in (ORACLE_ROOT / "schemas", ORACLE_ROOT / "corpus")
        for path in folder.rglob("*")
        if path.is_file()
    } | {
        "../../../packages/plugins/markdown/src/docwen_plugin_markdown/document_semantics_v3.py",
        "../../../packages/plugins/markdown/src/docwen_plugin_markdown/document_semantics_v3_fenced_source.py",
        "../../../packages/plugins/markdown/tests/test_document_semantics_v3.py",
        "../../../packages/plugins/markdown/tests/test_document_semantics_v3_anchor_topology.py",
        "../../../packages/plugins/markdown/tests/test_document_semantics_v3_fenced_source.py",
        "../../../tests/test_repo/test_markdown_semantics_v3_oracle.py",
    }
    for record in records:
        path = (ORACLE_ROOT / record["path"]).resolve()
        payload = path.read_bytes()
        assert path.is_file()
        assert record["bytes"] == len(payload)
        assert record["sha256"] == hashlib.sha256(payload).hexdigest()
    assert manifest["case_ids"] == [
        _json(path)["case_id"] for path in sorted((ORACLE_ROOT / "corpus").glob("*.case.json"))
    ]
    assert manifest["excluded"] == [
        "docwen.markdown_semantics.v1",
        "docwen.markdown_semantics.v2",
        "consumer.interop.v3",
        "attempt04",
    ]
