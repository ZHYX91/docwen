"""Typed, request-scoped context for optional conversion sidecar manifests."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, replace
from typing import Any


@dataclass(frozen=True, slots=True)
class OutputManifestPolicy:
    """Privacy-preserving projection of ``output.manifest`` configuration."""

    save_to_output: bool = False
    mask_input_path: bool = True

    @classmethod
    def from_config_snapshot(cls, snapshot: object) -> OutputManifestPolicy:
        root = snapshot if isinstance(snapshot, dict) else {}
        output = root.get("output")
        manifest = output.get("manifest") if isinstance(output, dict) else None
        values = manifest if isinstance(manifest, dict) else {}
        return cls(
            save_to_output=values.get("save_to_output") is True,
            # Privacy fails closed: only the literal boolean false disables masking.
            mask_input_path=values.get("mask_input_path") is not False,
        )

    def to_dict(self) -> dict[str, bool]:
        return {
            "save_to_output": self.save_to_output,
            "mask_input_path": self.mask_input_path,
        }

    @classmethod
    def from_dict(cls, data: object) -> OutputManifestPolicy:
        values = data if isinstance(data, dict) else {}
        return cls(
            save_to_output=values.get("save_to_output") is True,
            mask_input_path=values.get("mask_input_path") is not False,
        )


@dataclass(frozen=True, slots=True)
class ConversionManifestInput:
    """Original input identity frozen before any private pre-conversion."""

    path: str
    format: str
    category: str

    def to_dict(self) -> dict[str, str]:
        return {"path": self.path, "format": self.format, "category": self.category}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConversionManifestInput:
        return cls(
            path=str(data.get("path", "")),
            format=str(data.get("format", "")),
            category=str(data.get("category", "")),
        )


@dataclass(frozen=True, slots=True)
class PreconversionStep:
    """One bounded Office/pre-conversion step, without private staging paths."""

    input_index: int
    source_format: str
    target_format: str
    status: str
    backend: str = ""
    diagnostic_code: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "input_index": self.input_index,
            "source_format": self.source_format,
            "target_format": self.target_format,
            "status": self.status,
            "backend": self.backend,
            "diagnostic_code": self.diagnostic_code,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PreconversionStep:
        return cls(
            input_index=int(data.get("input_index", 0)),
            source_format=str(data.get("source_format", "")),
            target_format=str(data.get("target_format", "")),
            status=str(data.get("status", "")),
            backend=str(data.get("backend", "")),
            diagnostic_code=str(data.get("diagnostic_code", "")),
        )


@dataclass(frozen=True, slots=True)
class ConversionManifestContext:
    """Immutable manifest context carried with one conversion request."""

    policy: OutputManifestPolicy
    inputs: tuple[ConversionManifestInput, ...]
    preconversion_steps: tuple[PreconversionStep, ...] = ()
    batch_child: bool = False

    @classmethod
    def from_request_inputs(
        cls,
        input_refs: Iterable[object],
        config_snapshot: object,
    ) -> ConversionManifestContext:
        inputs = tuple(
            ConversionManifestInput(
                path=str(getattr(ref, "path", "")),
                format=str(getattr(ref, "format", "")),
                category=str(getattr(ref, "category", "")),
            )
            for ref in input_refs
            if str(getattr(ref, "input_role", "source")) in {"source", "neutral_document"}
        )
        return cls(policy=OutputManifestPolicy.from_config_snapshot(config_snapshot), inputs=inputs)

    def with_step(self, step: PreconversionStep) -> ConversionManifestContext:
        return replace(self, preconversion_steps=(*self.preconversion_steps, step))

    def for_input(self, index: int, *, batch_child: bool = True) -> ConversionManifestContext:
        selected_inputs = (self.inputs[index],) if 0 <= index < len(self.inputs) else ()
        selected_steps = tuple(
            replace(step, input_index=0) for step in self.preconversion_steps if step.input_index == index
        )
        return ConversionManifestContext(
            policy=self.policy,
            inputs=selected_inputs,
            preconversion_steps=selected_steps,
            batch_child=batch_child,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "policy": self.policy.to_dict(),
            "inputs": [item.to_dict() for item in self.inputs],
            "preconversion_steps": [step.to_dict() for step in self.preconversion_steps],
            "batch_child": self.batch_child,
        }

    @classmethod
    def from_dict(cls, data: object) -> ConversionManifestContext:
        values = data if isinstance(data, dict) else {}
        raw_inputs = values.get("inputs")
        raw_steps = values.get("preconversion_steps")
        return cls(
            policy=OutputManifestPolicy.from_dict(values.get("policy")),
            inputs=tuple(ConversionManifestInput.from_dict(item) for item in raw_inputs if isinstance(item, dict))
            if isinstance(raw_inputs, list)
            else (),
            preconversion_steps=tuple(PreconversionStep.from_dict(item) for item in raw_steps if isinstance(item, dict))
            if isinstance(raw_steps, list)
            else (),
            batch_child=values.get("batch_child") is True,
        )
