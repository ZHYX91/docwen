"""Packaged/source schema loading and validation for Machine Protocol v1."""

from __future__ import annotations

import json
from importlib import resources
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator
from referencing import Registry, Resource

_MACHINE_SCHEMA_NAME = "docwen.machine.v1.schema.json"
_DIAGNOSTIC_EVIDENCE_SCHEMA_NAME = "docwen.machine.diagnostic_evidence.v1.schema.json"
_BUNDLE_SCHEMA_NAME = "docwen.artifact_bundle.v2.schema.json"


class MachineContractValidator:
    """Validate every inbound and outbound wire object against DocWen schemas."""

    def __init__(self, contracts_root: Path | None = None) -> None:
        machine_schema = self._load_schema(_MACHINE_SCHEMA_NAME, contracts_root)
        diagnostic_evidence_schema = self._load_schema(_DIAGNOSTIC_EVIDENCE_SCHEMA_NAME, contracts_root)
        bundle_schema = self._load_schema(_BUNDLE_SCHEMA_NAME, contracts_root)
        registry = Registry().with_resources(
            [
                (bundle_schema["$id"], Resource.from_contents(bundle_schema)),
                (
                    diagnostic_evidence_schema["$id"],
                    Resource.from_contents(diagnostic_evidence_schema),
                ),
            ]
        )
        Draft202012Validator.check_schema(machine_schema)
        Draft202012Validator.check_schema(diagnostic_evidence_schema)
        Draft202012Validator.check_schema(bundle_schema)
        self._machine = Draft202012Validator(machine_schema, registry=registry)
        self._bundle = Draft202012Validator(bundle_schema, registry=registry)

    def validate_message(self, payload: dict[str, Any]) -> None:
        self._machine.validate(payload)

    def validate_bundle(self, payload: dict[str, Any]) -> None:
        self._bundle.validate(payload)

    @staticmethod
    def _load_schema(name: str, contracts_root: Path | None) -> dict[str, Any]:
        if contracts_root is not None:
            return json.loads((contracts_root / "schemas" / name).read_text(encoding="utf-8"))

        packaged = resources.files("docwen_cli").joinpath("contracts", name)
        if packaged.is_file():
            return json.loads(packaged.read_text(encoding="utf-8"))

        source_file = Path(__file__).resolve()
        for parent in source_file.parents:
            candidate = parent / "contracts" / "schemas" / name
            if candidate.is_file():
                return json.loads(candidate.read_text(encoding="utf-8"))
        raise FileNotFoundError(f"DocWen contract schema is unavailable: {name}")


__all__ = ["MachineContractValidator"]
